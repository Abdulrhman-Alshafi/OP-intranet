import * as React from 'react';
import { useCallback, useRef, useState } from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { PrimaryButton, DefaultButton, ActionButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { ProgressIndicator } from '@fluentui/react/lib/ProgressIndicator';
import { Icon } from '@fluentui/react/lib/Icon';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { MSGraphClientFactory, SPHttpClient } from '@microsoft/sp-http';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import * as XLSX from 'xlsx';

import * as strings from 'HrDocumentWebPartStrings';
import styles from './HrDocument.module.scss';
import { HrDocumentService, IHrUploadResult, IHrUploadJob } from '../../../services/HrDocumentService';

// ── Props ─────────────────────────────────────────────────────────────────────

export interface IHrUploadPanelProps {
  service: HrDocumentService;
  libraryName: string;
  hrGroupId: number;
  currentUserLoginName: string;
  siteUrl: string;
  spHttpClient: SPHttpClient;
  msGraphClientFactory: MSGraphClientFactory;
  onUploadsComplete: () => void;
}

// ── Types ─────────────────────────────────────────────────────────────────────

interface ISelectedEmployee {
  loginName: string;
  email: string;
  displayName: string;
  spId: number;
}

interface IExcelRow {
  fileName: string;
  employeeEmail: string;
  documentType: string;
}

interface IValidationReport {
  missingFiles: string[];
  extraFiles: string[];
  duplicateNames: string[];
  invalidEmails: string[];
  alreadyExistsFiles: string[];
}

// ── Document type options ─────────────────────────────────────────────────────

const DOCUMENT_TYPE_OPTIONS: IDropdownOption[] = [
  { key: 'Offer Letter',       text: 'Offer Letter' },
  { key: 'Contract',           text: 'Contract' },
  { key: 'Warning Letter',     text: 'Warning Letter' },
  { key: 'Promotion Letter',   text: 'Promotion Letter' },
  { key: 'Termination Letter', text: 'Termination Letter' },
  { key: 'NDA',                text: 'NDA' }
];

// ── Component ─────────────────────────────────────────────────────────────────

export const HrUploadPanel: React.FC<IHrUploadPanelProps> = (props) => {
  const {
    service,
    libraryName,
    hrGroupId,
    siteUrl,
    spHttpClient,
    msGraphClientFactory,
    onUploadsComplete
  } = props;

  const pickerContext = React.useMemo(
    () => ({ absoluteUrl: siteUrl, msGraphClientFactory, spHttpClient }),
    [siteUrl, msGraphClientFactory, spHttpClient]
  );

  // ── SINGLE UPLOAD ──────────────────────────────────────────────────────────

  const singleFileInputRef = useRef<HTMLInputElement>(null);
  const [singleFile, setSingleFile] = useState<File | undefined>(undefined);
  const [singleEmployee, setSingleEmployee] = useState<ISelectedEmployee | undefined>(undefined);
  const [singleDocType, setSingleDocType] = useState<string>('Offer Letter');
  const [singleUploading, setSingleUploading] = useState<boolean>(false);
  const [singleResult, setSingleResult] = useState<IHrUploadResult | undefined>(undefined);
  const [singleFileExists, setSingleFileExists] = useState<boolean>(false);
  const [singleCheckingExists, setSingleCheckingExists] = useState<boolean>(false);

  const handleSingleFileChange = async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file = e.target.files?.[0];
    setSingleFile(file);
    setSingleResult(undefined);
    setSingleFileExists(false);
    if (file) {
      setSingleCheckingExists(true);
      try {
        const exists = await service.fileExists(libraryName, file.name);
        setSingleFileExists(exists);
      } catch {
        setSingleFileExists(false);
      } finally {
        setSingleCheckingExists(false);
      }
    }
  };

  const handleSingleEmployeeChange = useCallback(
    async (items: { id: string; loginName: string; secondaryText?: string; text?: string }[]) => {
      if (items.length === 0) {
        setSingleEmployee(undefined);
        return;
      }
      const picked = items[0];
      const email = picked.secondaryText || picked.id;
      try {
        const resolved = await service.resolveUser(email);
        if (resolved) {
          setSingleEmployee({
            loginName: resolved.loginName,
            email,
            displayName: resolved.displayName,
            spId: resolved.id
          });
        }
      } catch {
        setSingleEmployee(undefined);
      }
    },
    [service]
  );

  const handleSingleUpload = async (): Promise<void> => {
    if (!singleFile || !singleEmployee || singleFileExists) return;
    setSingleUploading(true);
    setSingleResult(undefined);
    try {
      await service.uploadHrDocument(
        libraryName,
        singleFile,
        singleEmployee.loginName,
        singleEmployee.spId,
        hrGroupId,
        singleDocType
      );
      setSingleResult({ fileName: singleFile.name, status: 'success' });
      setSingleFile(undefined);
      setSingleEmployee(undefined);
      setSingleFileExists(false);
      if (singleFileInputRef.current) singleFileInputRef.current.value = '';
      onUploadsComplete();
    } catch (err) {
      setSingleResult({
        fileName: singleFile.name,
        status: 'error',
        message: err instanceof Error ? err.message : 'Upload failed.'
      });
    } finally {
      setSingleUploading(false);
    }
  };

  // ── BULK UPLOAD ────────────────────────────────────────────────────────────

  const excelInputRef = useRef<HTMLInputElement>(null);
  const docsInputRef = useRef<HTMLInputElement>(null);

  const [excelRows, setExcelRows] = useState<IExcelRow[]>([]);
  const [bulkFiles, setBulkFiles] = useState<File[]>([]);
  const [validation, setValidation] = useState<IValidationReport | undefined>(undefined);
  const [bulkJobs, setBulkJobs] = useState<IHrUploadJob[]>([]);
  const [uploadResults, setUploadResults] = useState<IHrUploadResult[]>([]);
  const [bulkUploading, setBulkUploading] = useState<boolean>(false);
  const [bulkValidating, setBulkValidating] = useState<boolean>(false);

  const handleExcelChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows: { FileName?: string; EmployeeEmail?: string; DocumentType?: string }[] =
          XLSX.utils.sheet_to_json(sheet);
        const parsed: IExcelRow[] = rows
          .filter((r) => r.FileName && r.EmployeeEmail)
          .map((r) => ({
            fileName: (r.FileName as string).toLowerCase().trim(),
            employeeEmail: (r.EmployeeEmail as string).toLowerCase().trim(),
            documentType: r.DocumentType ? (r.DocumentType as string).trim() : ''
          }));
        setExcelRows(parsed);
        setValidation(undefined);
        setBulkJobs([]);
        setUploadResults([]);
      } catch {
        setExcelRows([]);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleDocsChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const files = Array.from(e.target.files || []);
    setBulkFiles(files);
    setValidation(undefined);
    setBulkJobs([]);
    setUploadResults([]);
  };

  const handleValidate = useCallback(async () => {
    setBulkValidating(true);
    setValidation(undefined);
    setBulkJobs([]);

    const fileNameMap = new Map<string, File>();
    for (const f of bulkFiles) {
      fileNameMap.set(f.name.toLowerCase().trim(), f);
    }

    const excelFileNames = excelRows.map((r) => r.fileName);
    const uploadedFileNames = Array.from(fileNameMap.keys());

    // Duplicates in Excel
    const seen = new Set<string>();
    const duplicateNames: string[] = [];
    for (const name of excelFileNames) {
      if (seen.has(name)) duplicateNames.push(name);
      seen.add(name);
    }

    // Missing: in Excel but not in uploaded files
    const missingFiles = excelFileNames.filter(
      (name, idx) => excelFileNames.indexOf(name) === idx && !fileNameMap.has(name)
    );

    // Extra: uploaded but not in Excel
    const extraFiles = uploadedFileNames.filter((name) => excelFileNames.indexOf(name) === -1);

    // Resolve emails
    const invalidEmails: string[] = [];
    const resolvedJobs: IHrUploadJob[] = [];

    const uniqueRows = excelRows.filter(
      (row, idx) => excelRows.findIndex((r) => r.fileName === row.fileName) === idx
    );

    for (const row of uniqueRows) {
      const file = fileNameMap.get(row.fileName);
      if (!file) continue;

      const resolved = await service.resolveUser(row.employeeEmail).catch(() => undefined);
      if (!resolved) {
        invalidEmails.push(row.employeeEmail);
      } else {
        resolvedJobs.push({
          file,
          employeeLoginName: resolved.loginName,
          employeeId: resolved.id,
          documentType: row.documentType
        });
      }
    }

    // Check which files already exist
    const alreadyExistsFiles: string[] = [];
    for (const row of uniqueRows) {
      const file = fileNameMap.get(row.fileName);
      if (!file) continue;
      const exists = await service.fileExists(libraryName, file.name).catch(() => false);
      if (exists) alreadyExistsFiles.push(file.name);
    }

    setValidation({ missingFiles, extraFiles, duplicateNames, invalidEmails, alreadyExistsFiles });
    const filteredJobs = resolvedJobs.filter((j) => alreadyExistsFiles.indexOf(j.file.name) === -1);
    setBulkJobs(filteredJobs);
    setUploadResults(
      filteredJobs.map((j) => ({ fileName: j.file.name, status: 'pending' }))
    );
    setBulkValidating(false);
  }, [excelRows, bulkFiles, service, libraryName]);

  const hasValidationErrors = (v: IValidationReport): boolean =>
    v.missingFiles.length > 0 ||
    v.extraFiles.length > 0 ||
    v.duplicateNames.length > 0 ||
    v.invalidEmails.length > 0;

  const handleBulkUpload = async (): Promise<void> => {
    if (bulkJobs.length === 0) return;
    setBulkUploading(true);
    setUploadResults(bulkJobs.map((j) => ({ fileName: j.file.name, status: 'uploading' })));

    await service.uploadBulkHrDocuments(libraryName, bulkJobs, hrGroupId, (result) => {
      setUploadResults((prev) =>
        prev.map((r) => (r.fileName === result.fileName ? result : r))
      );
    });

    setBulkUploading(false);
    onUploadsComplete();
  };

  const handleRetryFailed = (): void => {
    const failedFileNames = new Set(
      uploadResults.filter((r) => r.status === 'error').map((r) => r.fileName)
    );
    const retryJobs = bulkJobs.filter((j) =>
      failedFileNames.has(j.file.name.toLowerCase().trim())
    );
    if (retryJobs.length === 0) return;
    setBulkJobs(retryJobs);
    setUploadResults(retryJobs.map((j) => ({ fileName: j.file.name, status: 'pending' })));
  };

  const handleDownloadErrorReport = (): void => {
    const errors = uploadResults.filter((r) => r.status === 'error');
    if (errors.length === 0) return;
    const csv = ['FileName,Error', ...errors.map((e) => `"${e.fileName}","${e.message || ''}"`)].join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'bulk-upload-errors.csv';
    a.click();
    URL.revokeObjectURL(url);
  };

  const canProceedBulk =
    validation !== undefined &&
    !hasValidationErrors(validation) &&
    bulkJobs.length > 0 &&
    !bulkUploading;

  const hasFailedUploads = uploadResults.some((r) => r.status === 'error');
  const allDone = uploadResults.length > 0 && uploadResults.every((r) => r.status === 'success' || r.status === 'error');
  const completedCount = uploadResults.filter((r) => r.status === 'success' || r.status === 'error').length;
  const progressPercent = uploadResults.length > 0 ? completedCount / uploadResults.length : 0;

  // ── Render ─────────────────────────────────────────────────────────────────

  return (
    <div className={styles.uploadColumns}>

      {/* ══ Left column: Single Upload ══════════════════════════════════════ */}
      <div className={styles.uploadColLeft}>
        <h3 className={styles.uploadSectionTitle}>
          <Icon iconName="Upload" />
          {strings.UploadSectionTitle}
        </h3>

        <Stack tokens={{ childrenGap: 14 }}>
          {/* File picker */}
          <div className={styles.uploadField}>
            <DefaultButton
              text={strings.SelectFileLabel}
              iconProps={{ iconName: 'Attach' }}
              onClick={() => singleFileInputRef.current?.click()}
            />
            {singleFile && (
              <div className={styles.fileNameDisplay}>{singleFile.name}</div>
            )}
            <input
              ref={singleFileInputRef}
              type="file"
              accept=".pdf,.doc,.docx,.xlsx,.xls,.ppt,.pptx,.png,.jpg,.jpeg,.zip"
              className={styles.hiddenFileInput}
              onChange={handleSingleFileChange}
            />
          </div>

          {/* People picker */}
          <div className={styles.uploadField}>
            <PeoplePicker
              context={pickerContext}
              titleText={strings.SelectEmployeeLabel}
              personSelectionLimit={1}
              required={true}
              principalTypes={[PrincipalType.User]}
              resolveDelay={300}
              ensureUser={true}
              onChange={handleSingleEmployeeChange as (items: unknown[]) => void}
            />
          </div>

          {/* Document type */}
          <Dropdown
            label={strings.DocumentTypeLabel}
            selectedKey={singleDocType}
            options={DOCUMENT_TYPE_OPTIONS}
            onChange={(_, o) => setSingleDocType(String(o?.key ?? 'Offer Letter'))}
            styles={{ root: { minWidth: 200 } }}
          />

          {singleCheckingExists && (
            <MessageBar messageBarType={MessageBarType.info} isMultiline={false}>
              Checking if file already exists…
            </MessageBar>
          )}
          {singleFileExists && !singleCheckingExists && (
            <MessageBar messageBarType={MessageBarType.blocked} isMultiline>
              {strings.ValidationFileExists}
            </MessageBar>
          )}
          {singleResult && (
            <MessageBar
              messageBarType={
                singleResult.status === 'success' ? MessageBarType.success : MessageBarType.error
              }
            >
              {singleResult.status === 'success'
                ? `"${singleResult.fileName}" uploaded successfully.`
                : singleResult.message}
            </MessageBar>
          )}

          <PrimaryButton
            text={singleUploading ? strings.UploadInProgressMessage : strings.UploadButtonLabel}
            disabled={!singleFile || !singleEmployee || singleUploading || singleFileExists || singleCheckingExists}
            onClick={handleSingleUpload}
            iconProps={singleUploading ? undefined : { iconName: 'Upload' }}
            onRenderIcon={singleUploading ? () => <Spinner size={SpinnerSize.xSmall} /> : undefined}
          />
        </Stack>
      </div>

      {/* ── Divider ─────────────────────────────────────────────────────── */}
      <div className={styles.colDivider} />

      {/* ══ Right column: Bulk Upload ════════════════════════════════════════ */}
      <div className={styles.uploadColRight}>
        <h3 className={styles.uploadSectionTitle}>
          <Icon iconName="BulkUpload" />
          {strings.BulkUploadSectionTitle}
        </h3>

        <Stack tokens={{ childrenGap: 14 }}>
          <div style={{ fontSize: 12, color: 'var(--neutralSecondary, #605e5c)' }}>
            {strings.ExcelColumnsHint}
          </div>

          {/* File selectors row */}
          <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
            <Stack tokens={{ childrenGap: 4 }}>
              <DefaultButton
                text={strings.SelectExcelLabel}
                iconProps={{ iconName: 'ExcelDocument' }}
                onClick={() => excelInputRef.current?.click()}
              />
              {excelRows.length > 0 && (
                <div className={styles.fileNameDisplay}>{excelRows.length} row(s) loaded</div>
              )}
              <input
                ref={excelInputRef}
                type="file"
                accept=".xlsx"
                className={styles.hiddenFileInput}
                onChange={handleExcelChange}
              />
            </Stack>

            <Stack tokens={{ childrenGap: 4 }}>
              <DefaultButton
                text={strings.SelectDocumentsLabel}
                iconProps={{ iconName: 'BulkUpload' }}
                onClick={() => docsInputRef.current?.click()}
              />
              {bulkFiles.length > 0 && (
                <div className={styles.fileNameDisplay}>{bulkFiles.length} file(s) selected</div>
              )}
              <input
                ref={docsInputRef}
                type="file"
                multiple
                className={styles.hiddenFileInput}
                onChange={handleDocsChange}
              />
            </Stack>

            <PrimaryButton
              text={bulkValidating ? 'Validating…' : 'Validate'}
              iconProps={{ iconName: 'CheckMark' }}
              disabled={excelRows.length === 0 || bulkFiles.length === 0 || bulkValidating}
              onClick={handleValidate}
              onRenderIcon={bulkValidating ? () => <Spinner size={SpinnerSize.xSmall} /> : undefined}
            />
          </Stack>

          {/* Validation report */}
          {validation && (
            <Stack tokens={{ childrenGap: 6 }} className={styles.validationReport}>
              {validation.missingFiles.length > 0 && (
                <MessageBar messageBarType={MessageBarType.error} isMultiline>
                  <strong>{strings.ValidationMissingFiles}</strong>
                  <br />
                  {validation.missingFiles.join(', ')}
                </MessageBar>
              )}
              {validation.extraFiles.length > 0 && (
                <MessageBar messageBarType={MessageBarType.warning} isMultiline>
                  <strong>{strings.ValidationExtraFiles}</strong>
                  <br />
                  {validation.extraFiles.join(', ')}
                </MessageBar>
              )}
              {validation.duplicateNames.length > 0 && (
                <MessageBar messageBarType={MessageBarType.error} isMultiline>
                  <strong>{strings.ValidationDuplicates}</strong>
                  <br />
                  {validation.duplicateNames.join(', ')}
                </MessageBar>
              )}
              {validation.invalidEmails.length > 0 && (
                <MessageBar messageBarType={MessageBarType.error} isMultiline>
                  <strong>{strings.ValidationInvalidEmails}</strong>
                  <br />
                  {validation.invalidEmails.join(', ')}
                </MessageBar>
              )}
              {validation.alreadyExistsFiles.length > 0 && (
                <MessageBar messageBarType={MessageBarType.warning} isMultiline>
                  <strong>{strings.ValidationBulkFileExists}</strong>
                  <br />
                  {validation.alreadyExistsFiles.join(', ')}
                </MessageBar>
              )}
              {!hasValidationErrors(validation) && (
                <MessageBar messageBarType={MessageBarType.success}>
                  Validation passed — {bulkJobs.length} document(s) ready to upload.
                </MessageBar>
              )}
            </Stack>
          )}

          {/* Progress table */}
          {uploadResults.length > 0 && (
            <Stack tokens={{ childrenGap: 8 }}>
              {bulkUploading && (
                <ProgressIndicator
                  label={strings.UploadInProgressMessage}
                  percentComplete={progressPercent}
                />
              )}
              <div className={styles.tableScrollWrapper}>
                <table className={styles.progressTable}>
                  <thead>
                    <tr>
                      <th>File</th>
                      <th>Status</th>
                      <th>Message</th>
                    </tr>
                  </thead>
                  <tbody>
                    {uploadResults.map((r) => (
                      <tr key={r.fileName}>
                        <td>{r.fileName}</td>
                        <td>
                          <span
                            className={
                              r.status === 'pending'
                                ? styles.statusPending
                                : r.status === 'uploading'
                                ? styles.statusUploading
                                : r.status === 'success'
                                ? styles.statusSuccess
                                : styles.statusError
                            }
                          >
                            {r.status === 'uploading' && (
                              <Spinner size={SpinnerSize.xSmall} />
                            )}
                            {r.status}
                          </span>
                        </td>
                        <td>{r.message || ''}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Stack>
          )}

          {/* Action buttons */}
          <div className={styles.actionRow}>
            <PrimaryButton
              text={strings.ProceedBulkLabel}
              iconProps={{ iconName: 'Upload' }}
              disabled={!canProceedBulk}
              onClick={handleBulkUpload}
            />
            {allDone && hasFailedUploads && (
              <DefaultButton
                text={strings.RetryFailedLabel}
                iconProps={{ iconName: 'Refresh' }}
                onClick={handleRetryFailed}
              />
            )}
            {allDone && hasFailedUploads && (
              <ActionButton
                text={strings.DownloadErrorReportLabel}
                iconProps={{ iconName: 'Download' }}
                onClick={handleDownloadErrorReport}
              />
            )}
          </div>
        </Stack>
      </div>

    </div>
  );
};

export default HrUploadPanel;
