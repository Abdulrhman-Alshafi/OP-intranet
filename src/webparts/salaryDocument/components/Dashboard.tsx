import * as React from 'react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

import * as strings from 'SalaryDocumentWebPartStrings';
import { ISalaryDocumentProps } from './ISalaryDocumentProps';
import { SalaryTable } from './SalaryTable';
import { UploadPanel } from './UploadPanel';
import { ISalaryDocument } from '../../../services/SalaryDocumentService';
import styles from './SalaryDocument.module.scss';

const PAGE_SIZE = 50;

const MONTH_FILTER_OPTIONS: IDropdownOption[] = [
  { key: 0, text: 'All Months' },
  { key: 1, text: 'January' }, { key: 2, text: 'February' }, { key: 3, text: 'March' },
  { key: 4, text: 'April' }, { key: 5, text: 'May' }, { key: 6, text: 'June' },
  { key: 7, text: 'July' }, { key: 8, text: 'August' }, { key: 9, text: 'September' },
  { key: 10, text: 'October' }, { key: 11, text: 'November' }, { key: 12, text: 'December' }
];

export const SalaryDocumentWebPartDashboard: React.FC<ISalaryDocumentProps> = (props) => {
  const {
    service,
    libraryName,
    accountantGroupName,
    currentUserEmail,
    currentUserLoginName,
    siteUrl,
    spHttpClient,
    msGraphClientFactory
  } = props;

  // ── Auth state ────────────────────────────────────────────────────────────
  const [isAccountant, setIsAccountant] = useState<boolean>(false);
  const [isSiteAdmin, setIsSiteAdmin] = useState<boolean>(false);
  const [accountantsGroupId, setAccountantsGroupId] = useState<number>(0);
  const [authLoading, setAuthLoading] = useState<boolean>(true);
  const [authError, setAuthError] = useState<string | undefined>(undefined);

  // ── Data state ────────────────────────────────────────────────────────────
  const [documents, setDocuments] = useState<ISalaryDocument[]>([]);
  const [nextLink, setNextLink] = useState<string | undefined>(undefined);
  const [dataLoading, setDataLoading] = useState<boolean>(false);
  const [dataError, setDataError] = useState<string | undefined>(undefined);
  const [searchQuery, setSearchQuery] = useState<string>('');

  // ── My salary state (accountant's own documents) ──────────────────────────
  const [mySalaryDocuments, setMySalaryDocuments] = useState<ISalaryDocument[]>([]);
  const [mySearchQuery, setMySearchQuery] = useState<string>('');
  const [myDataLoading, setMyDataLoading] = useState<boolean>(false);
  const [myDataError, setMyDataError] = useState<string | undefined>(undefined);

  // ── Period filter state ──────────────────────────────────────────────────
  const [filterYear, setFilterYear] = useState<number>(0);
  const [filterMonth, setFilterMonth] = useState<number>(0);
  const [myFilterYear, setMyFilterYear] = useState<number>(0);
  const [myFilterMonth, setMyFilterMonth] = useState<number>(0);

  // Track whether still mounted to avoid state updates after unmount
  const mountedRef = useRef(true);

  // ── Auth check on mount ───────────────────────────────────────────────────
  useEffect(() => {
    mountedRef.current = true;
    let cancelled = false;

    (async () => {
      setAuthLoading(true);
      setAuthError(undefined);
      try {
        const [accountant, grpId, siteAdmin] = await Promise.all([
          service.isUserInGroup(accountantGroupName),
          service.resolveAccountantsGroupId(accountantGroupName),
          service.isCurrentUserSiteAdmin()
        ]);
        if (!cancelled) {
          setIsAccountant(accountant);
          setAccountantsGroupId(grpId);
          setIsSiteAdmin(siteAdmin);
        }
      } catch (err) {
        if (!cancelled) {
          setAuthError(
            err instanceof Error ? err.message : 'Unable to determine user role.'
          );
        }
      } finally {
        if (!cancelled) setAuthLoading(false);
      }
    })().catch(() => undefined);

    return () => {
      cancelled = true;
      mountedRef.current = false;
    };
  }, [service, accountantGroupName]);

  // ── Load documents once auth is resolved ─────────────────────────────────
  useEffect(() => {
    if (authLoading) return;
    let cancelled = false;

    (async () => {
      setDataLoading(true);
      setDataError(undefined);
      try {
        if (isAccountant) {
          const [page, myItems] = await Promise.all([
            service.getAllSalaryDocuments(libraryName, PAGE_SIZE),
            service.getSalaryDocuments(libraryName, currentUserEmail)
          ]);
          if (!cancelled) {
            setDocuments(page.items);
            setNextLink(page.nextLink);
            setMySalaryDocuments(myItems);
          }
        } else {
          const items = await service.getSalaryDocuments(libraryName, currentUserEmail);
          if (!cancelled) {
            setDocuments(items);
            setNextLink(undefined);
          }
        }
      } catch (err) {
        if (!cancelled) {
          setDataError(err instanceof Error ? err.message : 'Unable to load documents.');
        }
      } finally {
        if (!cancelled) setDataLoading(false);
      }
    })().catch(() => undefined);

    return () => {
      cancelled = true;
    };
  }, [authLoading, isAccountant, service, libraryName, currentUserEmail]);

  // ── Load next page (accountant pagination) ────────────────────────────────
  const handleLoadMore = useCallback(async () => {
    if (!nextLink) return;
    setDataLoading(true);
    try {
      const page = await service.getAllSalaryDocuments(libraryName, PAGE_SIZE, nextLink);
      if (mountedRef.current) {
        setDocuments((prev) => [...prev, ...page.items]);
        setNextLink(page.nextLink);
      }
    } catch (err) {
      if (mountedRef.current) {
        setDataError(err instanceof Error ? err.message : 'Failed to load more documents.');
      }
    } finally {
      if (mountedRef.current) setDataLoading(false);
    }
  }, [service, libraryName, nextLink]);

  // ── After a successful upload, reload data ────────────────────────────────
  const handleUploadsComplete = useCallback(async () => {
    setDataLoading(true);
    setMyDataLoading(true);
    setDataError(undefined);
    setMyDataError(undefined);
    try {
      const [page, myItems] = await Promise.all([
        service.getAllSalaryDocuments(libraryName, PAGE_SIZE),
        service.getSalaryDocuments(libraryName, currentUserEmail)
      ]);
      if (mountedRef.current) {
        setDocuments(page.items);
        setNextLink(page.nextLink);
        setMySalaryDocuments(myItems);
      }
    } catch (err) {
      if (mountedRef.current) {
        setDataError(err instanceof Error ? err.message : 'Failed to reload documents.');
      }
    } finally {
      if (mountedRef.current) {
        setDataLoading(false);
        setMyDataLoading(false);
      }
    }
  }, [service, libraryName, currentUserEmail]);

  // ── Delete a document ─────────────────────────────────────────────────────
  const handleDeleteDocument = useCallback(async (item: ISalaryDocument) => {
    await service.deleteDocument(libraryName, item.listItemId);
    if (mountedRef.current) {
      setDocuments((prev) => prev.filter((d) => d.listItemId !== item.listItemId));
      setMySalaryDocuments((prev) => prev.filter((d) => d.listItemId !== item.listItemId));
    }
  }, [service, libraryName]);

  // ── Client-side filtering ─────────────────────────────────────────────────
  const filteredDocuments = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    return documents.filter((doc) => {
      if (filterYear > 0 && doc.payPeriodYear !== filterYear) return false;
      if (filterMonth > 0 && doc.payPeriodMonth !== filterMonth) return false;
      if (!q) return true;
      const nameMatch = doc.name.toLowerCase().includes(q);
      const employeeMatch = isAccountant && doc.employeeDisplayName.toLowerCase().includes(q);
      return nameMatch || employeeMatch;
    });
  }, [documents, searchQuery, isAccountant, filterYear, filterMonth]);

  const filteredMyDocuments = useMemo(() => {
    const q = mySearchQuery.trim().toLowerCase();
    return mySalaryDocuments.filter((doc) => {
      if (myFilterYear > 0 && doc.payPeriodYear !== myFilterYear) return false;
      if (myFilterMonth > 0 && doc.payPeriodMonth !== myFilterMonth) return false;
      if (!q) return true;
      return doc.name.toLowerCase().includes(q);
    });
  }, [mySalaryDocuments, mySearchQuery, myFilterYear, myFilterMonth]);

  const yearOptions = useMemo((): IDropdownOption[] => {
    const years = new Set<number>();
    documents.forEach((d) => { if (d.payPeriodYear > 0) years.add(d.payPeriodYear); });
    const opts: IDropdownOption[] = [{ key: 0, text: strings.FilterAllYears }];
    Array.from(years).sort((a, b) => b - a).forEach((y) => opts.push({ key: y, text: String(y) }));
    return opts;
  }, [documents]);

  const myYearOptions = useMemo((): IDropdownOption[] => {
    const years = new Set<number>();
    mySalaryDocuments.forEach((d) => { if (d.payPeriodYear > 0) years.add(d.payPeriodYear); });
    const opts: IDropdownOption[] = [{ key: 0, text: strings.FilterAllYears }];
    Array.from(years).sort((a, b) => b - a).forEach((y) => opts.push({ key: y, text: String(y) }));
    return opts;
  }, [mySalaryDocuments]);

  // ── Render ────────────────────────────────────────────────────────────────
  if (authLoading) {
    return (
      <div className={styles.salaryDocument}>
        <div className={styles.wpHeader}>
          <div className={styles.wpTitleSection}>
            <h2 className={styles.wpTitle}>{strings.WebPartTitle}</h2>
          </div>
        </div>
        <Stack horizontalAlign="center" styles={{ root: { padding: 32 } }}>
          <Spinner size={SpinnerSize.large} />
        </Stack>
      </div>
    );
  }

  if (authError) {
    return (
      <div className={styles.salaryDocument}>
        <div className={styles.wpHeader}>
          <div className={styles.wpTitleSection}>
            <h2 className={styles.wpTitle}>{strings.WebPartTitle}</h2>
          </div>
        </div>
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
          {authError}
        </MessageBar>
      </div>
    );
  }

  const searchPlaceholder = isAccountant
    ? strings.SearchPlaceholderAccountant
    : strings.SearchPlaceholderEmployee;

  const tableContent = (
    <Stack tokens={{ childrenGap: 8 }}>
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 12, alignItems: 'center' }}>
        <Dropdown
          selectedKey={filterYear}
          options={yearOptions}
          onChange={(_, o) => setFilterYear(Number(o?.key ?? 0))}
          styles={{ root: { minWidth: 120, flex: '1 1 120px', maxWidth: 200 }, title: { height: 32, lineHeight: 30 } }}
        />
        <Dropdown
          selectedKey={filterMonth}
          options={MONTH_FILTER_OPTIONS}
          onChange={(_, o) => setFilterMonth(Number(o?.key ?? 0))}
          styles={{ root: { minWidth: 120, flex: '1 1 120px', maxWidth: 200 }, title: { height: 32, lineHeight: 30 } }}
        />
        <SearchBox
          placeholder={searchPlaceholder}
          value={searchQuery}
          onChange={(_, newValue) => setSearchQuery(newValue || '')}
          styles={{ root: { flex: '2 1 180px', minWidth: 180, height: 32, padding: 0 } }}
        />
      </div>
      {dataError && (
        <MessageBar messageBarType={MessageBarType.error}>
          {dataError}
        </MessageBar>
      )}
      <SalaryTable
        items={filteredDocuments}
        showEmployeeColumn={isAccountant}
        loading={dataLoading}
        nextLink={nextLink}
        onLoadMore={handleLoadMore}
        canDelete={isAccountant || isSiteAdmin}
        onDeleteRequest={handleDeleteDocument}
      />
    </Stack>
  );

  if (!isAccountant) {
    return (
      <div className={styles.salaryDocument}>
        <div className={styles.wpHeader}>
          <div className={styles.wpTitleSection}>
            <h2 className={styles.wpTitle}>{strings.WebPartTitle}</h2>
          </div>
        </div>
        {tableContent}
      </div>
    );
  }

  // Accountant: three-tab view (My Salary / All Documents / Upload)
  const mySalaryContent = (
    <Stack tokens={{ childrenGap: 8 }}>
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 12, alignItems: 'center' }}>
        <Dropdown
          selectedKey={myFilterYear}
          options={myYearOptions}
          onChange={(_, o) => setMyFilterYear(Number(o?.key ?? 0))}
          styles={{ root: { minWidth: 120, flex: '1 1 120px', maxWidth: 200 }, title: { height: 32, lineHeight: 30 } }}
        />
        <Dropdown
          selectedKey={myFilterMonth}
          options={MONTH_FILTER_OPTIONS}
          onChange={(_, o) => setMyFilterMonth(Number(o?.key ?? 0))}
          styles={{ root: { minWidth: 120, flex: '1 1 120px', maxWidth: 200 }, title: { height: 32, lineHeight: 30 } }}
        />
        <SearchBox
          placeholder={strings.SearchPlaceholderEmployee}
          value={mySearchQuery}
          onChange={(_, newValue) => setMySearchQuery(newValue || '')}
          styles={{ root: { flex: '2 1 180px', minWidth: 180, height: 32, padding: 0 } }}
        />
      </div>
      {myDataError && (
        <MessageBar messageBarType={MessageBarType.error}>
          {myDataError}
        </MessageBar>
      )}
      <SalaryTable
        items={filteredMyDocuments}
        showEmployeeColumn={false}
        loading={myDataLoading}
        nextLink={undefined}
        onLoadMore={async () => undefined}
        canDelete={isAccountant || isSiteAdmin}
        onDeleteRequest={handleDeleteDocument}
      />
    </Stack>
  );

  return (
    <div className={styles.salaryDocument}>
      <div className={styles.wpHeader}>
        <div className={styles.wpTitleSection}>
          <h2 className={styles.wpTitle}>{strings.WebPartTitle}</h2>
        </div>
      </div>
      <Pivot>
        <PivotItem headerText={strings.TabMySalary}>
          <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 16 } }}>
            {mySalaryContent}
          </Stack>
        </PivotItem>
        <PivotItem headerText={strings.TabDocuments}>
          <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 16 } }}>
            {tableContent}
          </Stack>
        </PivotItem>
        <PivotItem headerText={strings.TabUpload}>
          <div style={{ marginTop: 16 }}>
            <UploadPanel
              service={service}
              libraryName={libraryName}
              accountantsGroupId={accountantsGroupId}
              currentUserLoginName={currentUserLoginName}
              siteUrl={siteUrl}
              spHttpClient={spHttpClient}
              msGraphClientFactory={msGraphClientFactory}
              onUploadsComplete={handleUploadsComplete}
            />
          </div>
        </PivotItem>
      </Pivot>
    </div>
  );
};

export default SalaryDocumentWebPartDashboard;
