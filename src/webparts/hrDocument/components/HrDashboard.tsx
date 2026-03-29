import * as React from 'react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

import * as strings from 'HrDocumentWebPartStrings';
import { IHrDocumentProps } from './IHrDocumentProps';
import { HrTable } from './HrTable';
import { HrUploadPanel } from './HrUploadPanel';
import { IHrDocument } from '../../../services/HrDocumentService';
import styles from './HrDocument.module.scss';

const PAGE_SIZE = 50;

export const HrDashboard: React.FC<IHrDocumentProps> = (props) => {
  const {
    service,
    libraryName,
    hrGroupName,
    currentUserEmail,
    currentUserLoginName,
    siteUrl,
    spHttpClient,
    msGraphClientFactory
  } = props;

  // ── Auth state ────────────────────────────────────────────────────────────
  const [isHr, setIsHr] = useState<boolean>(false);
  const [isSiteAdmin, setIsSiteAdmin] = useState<boolean>(false);
  const [hrGroupId, setHrGroupId] = useState<number>(0);
  const [authLoading, setAuthLoading] = useState<boolean>(true);
  const [authError, setAuthError] = useState<string | undefined>(undefined);

  // ── Data state (HR: all documents) ───────────────────────────────────────
  const [documents, setDocuments] = useState<IHrDocument[]>([]);
  const [nextLink, setNextLink] = useState<string | undefined>(undefined);
  const [dataLoading, setDataLoading] = useState<boolean>(false);
  const [dataError, setDataError] = useState<string | undefined>(undefined);
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [filterDocType, setFilterDocType] = useState<string>('');

  // ── My documents state (own HR letters) ──────────────────────────────────
  const [myDocuments, setMyDocuments] = useState<IHrDocument[]>([]);
  const [mySearchQuery, setMySearchQuery] = useState<string>('');
  const [myDataLoading, setMyDataLoading] = useState<boolean>(false);
  const [myDataError, setMyDataError] = useState<string | undefined>(undefined);
  const [myFilterDocType, setMyFilterDocType] = useState<string>('');

  const mountedRef = useRef(true);

  // ── Auth check on mount ───────────────────────────────────────────────────
  useEffect(() => {
    mountedRef.current = true;
    let cancelled = false;

    (async () => {
      setAuthLoading(true);
      setAuthError(undefined);
      try {
        const [hrMember, grpId, siteAdmin] = await Promise.all([
          service.isUserInGroup(hrGroupName),
          service.resolveHrGroupId(hrGroupName),
          service.isCurrentUserSiteAdmin()
        ]);
        if (!cancelled) {
          setIsHr(hrMember);
          setHrGroupId(grpId);
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
  }, [service, hrGroupName]);

  // ── Load documents once auth is resolved ─────────────────────────────────
  useEffect(() => {
    if (authLoading) return;
    let cancelled = false;

    (async () => {
      setDataLoading(true);
      setDataError(undefined);
      try {
        if (isHr || isSiteAdmin) {
          const [page, myItems] = await Promise.all([
            service.getAllHrDocuments(libraryName, PAGE_SIZE),
            service.getHrDocuments(libraryName, currentUserEmail)
          ]);
          if (!cancelled) {
            setDocuments(page.items);
            setNextLink(page.nextLink);
            setMyDocuments(myItems);
          }
        } else {
          const items = await service.getHrDocuments(libraryName, currentUserEmail);
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

    return () => { cancelled = true; };
  }, [authLoading, isHr, isSiteAdmin, service, libraryName, currentUserEmail]);

  // ── Load next page ────────────────────────────────────────────────────────
  const handleLoadMore = useCallback(async () => {
    if (!nextLink) return;
    setDataLoading(true);
    try {
      const page = await service.getAllHrDocuments(libraryName, PAGE_SIZE, nextLink);
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
        service.getAllHrDocuments(libraryName, PAGE_SIZE),
        service.getHrDocuments(libraryName, currentUserEmail)
      ]);
      if (mountedRef.current) {
        setDocuments(page.items);
        setNextLink(page.nextLink);
        setMyDocuments(myItems);
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
  const handleDeleteDocument = useCallback(async (item: IHrDocument) => {
    await service.deleteDocument(libraryName, item.listItemId);
    if (mountedRef.current) {
      setDocuments((prev) => prev.filter((d) => d.listItemId !== item.listItemId));
      setMyDocuments((prev) => prev.filter((d) => d.listItemId !== item.listItemId));
    }
  }, [service, libraryName]);

  // ── Client-side filtering ─────────────────────────────────────────────────
  const filteredDocuments = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    return documents.filter((doc) => {
      if (filterDocType && doc.documentType !== filterDocType) return false;
      if (!q) return true;
      const nameMatch = doc.name.toLowerCase().indexOf(q) !== -1;
      const empMatch = (isHr || isSiteAdmin) && doc.employeeDisplayName.toLowerCase().indexOf(q) !== -1;
      return nameMatch || empMatch;
    });
  }, [documents, searchQuery, isHr, isSiteAdmin, filterDocType]);

  const filteredMyDocuments = useMemo(() => {
    const q = mySearchQuery.trim().toLowerCase();
    return myDocuments.filter((doc) => {
      if (myFilterDocType && doc.documentType !== myFilterDocType) return false;
      if (!q) return true;
      return doc.name.toLowerCase().indexOf(q) !== -1;
    });
  }, [myDocuments, mySearchQuery, myFilterDocType]);

  // ── Document type dropdown options (dynamic from loaded docs) ─────────────
  const docTypeOptions = useMemo((): IDropdownOption[] => {
    const types = new Set<string>();
    documents.forEach((d) => { if (d.documentType) types.add(d.documentType); });
    const opts: IDropdownOption[] = [{ key: '', text: strings.FilterAllTypes }];
    Array.from(types).sort().forEach((t) => opts.push({ key: t, text: t }));
    return opts;
  }, [documents]);

  const myDocTypeOptions = useMemo((): IDropdownOption[] => {
    const types = new Set<string>();
    myDocuments.forEach((d) => { if (d.documentType) types.add(d.documentType); });
    const opts: IDropdownOption[] = [{ key: '', text: strings.FilterAllTypes }];
    Array.from(types).sort().forEach((t) => opts.push({ key: t, text: t }));
    return opts;
  }, [myDocuments]);

  // ── Render ────────────────────────────────────────────────────────────────
  if (authLoading) {
    return (
      <div className={styles.hrDocument}>
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
      <div className={styles.hrDocument}>
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

  const isManager = isHr || isSiteAdmin;
  const searchPlaceholder = isManager
    ? strings.SearchPlaceholderHr
    : strings.SearchPlaceholderEmployee;

  // ── All Documents view (used in both employee tab and HR "Documents" tab)
  const allDocsContent = (
    <Stack tokens={{ childrenGap: 8 }}>
      <Stack
        horizontal
        wrap
        tokens={{ childrenGap: 16 }}
        styles={{ root: { rowGap: 12, alignItems: 'stretch' } }}
      >
        <Dropdown
          selectedKey={filterDocType}
          options={docTypeOptions}
          onChange={(_, o) => setFilterDocType(String(o?.key ?? ''))}
          styles={{ root: { minWidth: 160, margin: 0 }, title: { height: 32, lineHeight: 30 } }}
        />
        <SearchBox
          placeholder={searchPlaceholder}
          value={searchQuery}
          onChange={(_, newValue) => setSearchQuery(newValue || '')}
          styles={{ root: { flexGrow: 1, minWidth: 180, height: 32, margin: 0, padding: 0 } }}
        />
      </Stack>
      {dataError && (
        <MessageBar messageBarType={MessageBarType.error}>{dataError}</MessageBar>
      )}
      <HrTable
        items={filteredDocuments}
        showEmployeeColumn={isManager}
        loading={dataLoading}
        nextLink={nextLink}
        onLoadMore={handleLoadMore}
        canDelete={isManager}
        onDeleteRequest={handleDeleteDocument}
      />
    </Stack>
  );

  // ── Employee view: just the table, no tabs ────────────────────────────────
  if (!isManager) {
    return (
      <div className={styles.hrDocument}>
        <div className={styles.wpHeader}>
          <div className={styles.wpTitleSection}>
            <h2 className={styles.wpTitle}>{strings.WebPartTitle}</h2>
          </div>
        </div>
        {allDocsContent}
      </div>
    );
  }

  // ── HR / Admin: three-tab view ────────────────────────────────────────────
  const myDocsContent = (
    <Stack tokens={{ childrenGap: 8 }}>
      <Stack
        horizontal
        wrap
        tokens={{ childrenGap: 16 }}
        styles={{ root: { rowGap: 12, alignItems: 'stretch' } }}
      >
        <Dropdown
          selectedKey={myFilterDocType}
          options={myDocTypeOptions}
          onChange={(_, o) => setMyFilterDocType(String(o?.key ?? ''))}
          styles={{ root: { minWidth: 160, margin: 0 }, title: { height: 32, lineHeight: 30 } }}
        />
        <SearchBox
          placeholder={strings.SearchPlaceholderEmployee}
          value={mySearchQuery}
          onChange={(_, newValue) => setMySearchQuery(newValue || '')}
          styles={{ root: { flexGrow: 1, minWidth: 180, height: 32, margin: 0, padding: 0 } }}
        />
      </Stack>
      {myDataError && (
        <MessageBar messageBarType={MessageBarType.error}>{myDataError}</MessageBar>
      )}
      <HrTable
        items={filteredMyDocuments}
        showEmployeeColumn={false}
        loading={myDataLoading}
        nextLink={undefined}
        onLoadMore={async () => undefined}
        canDelete={isManager}
        onDeleteRequest={handleDeleteDocument}
      />
    </Stack>
  );

  return (
    <div className={styles.hrDocument}>
      <div className={styles.wpHeader}>
        <div className={styles.wpTitleSection}>
          <h2 className={styles.wpTitle}>{strings.WebPartTitle}</h2>
        </div>
      </div>
      <Pivot>
        <PivotItem headerText={strings.TabMyDocuments}>
          <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 16 } }}>
            {myDocsContent}
          </Stack>
        </PivotItem>
        <PivotItem headerText={strings.TabDocuments}>
          <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 16 } }}>
            {allDocsContent}
          </Stack>
        </PivotItem>
        <PivotItem headerText={strings.TabUpload}>
          <div style={{ marginTop: 16 }}>
            <HrUploadPanel
              service={service}
              libraryName={libraryName}
              hrGroupId={hrGroupId}
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

export default HrDashboard;
