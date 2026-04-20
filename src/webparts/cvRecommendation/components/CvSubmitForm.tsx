import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import styles from './CvRecommendation.module.scss';

export interface ICvSubmitFormProps {
  isOpen: boolean;
  onDismiss: () => void;
  onSubmit: (
    candidateName: string,
    candidateEmail: string,
    phoneNumber: string,
    position: string,
    notes: string,
    file?: File
  ) => Promise<void>;
}

const CvSubmitForm: React.FC<ICvSubmitFormProps> = ({ isOpen, onDismiss, onSubmit }) => {
  const [candidateName, setCandidateName] = React.useState('');
  const [candidateEmail, setCandidateEmail] = React.useState('');
  const [phoneNumber, setPhoneNumber] = React.useState('');
  const [position, setPosition] = React.useState('');
  const [notes, setNotes] = React.useState('');
  const [file, setFile] = React.useState<File | undefined>(undefined);
  const [saving, setSaving] = React.useState(false);
  const [error, setError] = React.useState('');
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const reset = React.useCallback((): void => {
    setCandidateName('');
    setCandidateEmail('');
    setPhoneNumber('');
    setPosition('');
    setNotes('');
    setFile(undefined);
    setError('');
  }, []);

  const validate = (): string => {
    if (!candidateName.trim()) return 'Candidate Name is required.';
    if (!candidateEmail.trim()) return 'Candidate Email is required.';
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(candidateEmail.trim())) return 'Enter a valid email address.';
    if (!position.trim()) return 'Position is required.';
    return '';
  };

  const handleSubmit = React.useCallback(async (): Promise<void> => {
    const err = validate();
    if (err) { setError(err); return; }

    setSaving(true);
    setError('');
    try {
      await onSubmit(
        candidateName.trim(),
        candidateEmail.trim(),
        phoneNumber.trim(),
        position.trim(),
        notes.trim(),
        file
      );
      reset();
      onDismiss();
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Failed to submit. Please try again.');
    } finally {
      setSaving(false);
    }
  }, [candidateName, candidateEmail, phoneNumber, position, notes, file, onSubmit, onDismiss, reset]);

  const handleDismiss = React.useCallback((): void => {
    if (!saving) { reset(); onDismiss(); }
  }, [saving, reset, onDismiss]);

  const handleFileChange = React.useCallback((e: React.ChangeEvent<HTMLInputElement>): void => {
    const selected = e.target.files?.[0];
    if (selected) setFile(selected);
  }, []);

  const handleRemoveFile = React.useCallback((): void => {
    setFile(undefined);
    if (fileInputRef.current) fileInputRef.current.value = '';
  }, []);

  const onRenderFooter = React.useCallback((): JSX.Element => (
    <Stack horizontal tokens={{ childrenGap: 8 }}>
      <PrimaryButton onClick={handleSubmit} disabled={saving}>
        {saving ? <Spinner size={SpinnerSize.xSmall} style={{ marginRight: 6 }} /> : null}
        {saving ? 'Submitting…' : 'Submit CV'}
      </PrimaryButton>
      <DefaultButton onClick={handleDismiss} disabled={saving}>Cancel</DefaultButton>
    </Stack>
  ), [handleSubmit, handleDismiss, saving]);

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={handleDismiss}
      type={PanelType.medium}
      headerText="Recommend a Candidate CV"
      onRenderFooterContent={onRenderFooter}
      isFooterAtBottom
      closeButtonAriaLabel="Close"
    >
      <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 16 }}>

        {error && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError('')}>
            {error}
          </MessageBar>
        )}

        {/* Candidate Details */}
        <div className={styles.formSection}>
          <p className={styles.formSectionTitle}>Candidate Details</p>
          <Stack tokens={{ childrenGap: 12 }}>
            <TextField
              label="Candidate Name"
              required
              value={candidateName}
              onChange={(_, v) => setCandidateName(v ?? '')}
              placeholder="Full name"
              disabled={saving}
            />
            <TextField
              label="Candidate Email"
              required
              type="email"
              value={candidateEmail}
              onChange={(_, v) => setCandidateEmail(v ?? '')}
              placeholder="email@example.com"
              disabled={saving}
            />
            <TextField
              label="Phone Number"
              value={phoneNumber}
              onChange={(_, v) => setPhoneNumber(v ?? '')}
              placeholder="+966 5X XXX XXXX"
              disabled={saving}
            />
          </Stack>
        </div>

        {/* Role */}
        <div className={styles.formSection}>
          <p className={styles.formSectionTitle}>Application</p>
          <Stack tokens={{ childrenGap: 12 }}>
            <TextField
              label="Position Applied For"
              required
              value={position}
              onChange={(_, v) => setPosition(v ?? '')}
              placeholder="e.g. Senior Software Engineer"
              disabled={saving}
            />
            <TextField
              label="Notes / Comments"
              multiline
              rows={4}
              value={notes}
              onChange={(_, v) => setNotes(v ?? '')}
              placeholder="Add any notes or comments about this candidate…"
              disabled={saving}
            />
          </Stack>
        </div>

        {/* CV File */}
        <div className={styles.formSection}>
          <p className={styles.formSectionTitle}>CV Attachment</p>
          {file ? (
            <div className={styles.fileSelected}>
              <Icon iconName="PDF" className={styles.fileIcon} />
              <span className={styles.fileName}>{file.name}</span>
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                title="Remove file"
                onClick={handleRemoveFile}
                disabled={saving}
              />
            </div>
          ) : (
            <div
              className={styles.fileUploadArea}
              onClick={() => !saving && fileInputRef.current?.click()}
              role="button"
              tabIndex={0}
              onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') fileInputRef.current?.click(); }}
            >
              <div className={styles.fileUploadIcon}>
                <Icon iconName="Upload" />
              </div>
              <div className={styles.fileUploadText}>Click to upload CV file</div>
              <div className={styles.fileUploadSubtext}>PDF, DOC, DOCX up to 10 MB</div>
            </div>
          )}
          <input
            ref={fileInputRef}
            type="file"
            accept=".pdf,.doc,.docx"
            style={{ display: 'none' }}
            onChange={handleFileChange}
          />
        </div>

      </Stack>
    </Panel>
  );
};

export default CvSubmitForm;
