import * as React from 'react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

import * as strings from 'SalaryDocumentWebPartStrings';
import { ISalaryDocumentProps } from './ISalaryDocumentProps';
import { SalaryTable } from './SalaryTable';
import { UploadPanel } from './UploadPanel';
import { ISalaryDocument, IPagedSalaryDocuments } from '../../../services/SalaryDocumentService';

const PAGE_SIZE = 50;

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
  const [accountantsGroupId, setAccountantsGroupId] = useState<number>(0);
  const [authLoading, setAuthLoading] = useState<boolean>(true);
  const [authError, setAuthError] = useState<string | undefined>(undefined);

  // ── Data state ────────────────────────────────────────────────────────────
  const [documents, setDocuments] = useState<ISalaryDocument[]>([]);
  const [nextLink, setNextLink] = useState<string | undefined>(undefined);
  const [dataLoading, setDataLoading] = useState<boolean>(false);
  const [dataError, setDataError] = useState<string | undefined>(undefined);
  const [searchQuery, setSearchQuery] = useState<string>('');

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
        const [accountant, grpId] = await Promise.all([
          service.isUserInGroup(accountantGroupName),
          service.resolveAccountantsGroupId(accountantGroupName)
        ]);
        if (!cancelled) {
          setIsAccountant(accountant);
          setAccountantsGroupId(grpId);
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
          const page: IPagedSalaryDocuments = await service.getAllSalaryDocuments(
            libraryName,
            PAGE_SIZE
          );
          if (!cancelled) {
            setDocuments(page.items);
            setNextLink(page.nextLink);
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
    setDataError(undefined);
    try {
      const page = await service.getAllSalaryDocuments(libraryName, PAGE_SIZE);
      if (mountedRef.current) {
        setDocuments(page.items);
        setNextLink(page.nextLink);
      }
    } catch (err) {
      if (mountedRef.current) {
        setDataError(err instanceof Error ? err.message : 'Failed to reload documents.');
      }
    } finally {
      if (mountedRef.current) setDataLoading(false);
    }
  }, [service, libraryName]);

  // ── Client-side filtering ─────────────────────────────────────────────────
  const filteredDocuments = useMemo(() => {
    const q = searchQuery.trim().toLowerCase();
    if (!q) return documents;
    return documents.filter((doc) => {
      const nameMatch = doc.name.toLowerCase().includes(q);
      const employeeMatch = isAccountant && doc.employeeDisplayName.toLowerCase().includes(q);
      return nameMatch || employeeMatch;
    });
  }, [documents, searchQuery, isAccountant]);

  // ── Render ────────────────────────────────────────────────────────────────
  if (authLoading) {
    return (
      <Stack horizontalAlign="center" styles={{ root: { padding: 32 } }}>
        <Spinner size={SpinnerSize.large} />
      </Stack>
    );
  }

  if (authError) {
    return (
      <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
        {authError}
      </MessageBar>
    );
  }

  const searchPlaceholder = isAccountant
    ? strings.SearchPlaceholderAccountant
    : strings.SearchPlaceholderEmployee;

  const tableContent = (
    <Stack tokens={{ childrenGap: 8 }}>
      <SearchBox
        placeholder={searchPlaceholder}
        value={searchQuery}
        onChange={(_, newValue) => setSearchQuery(newValue || '')}
        styles={{ root: { maxWidth: 400 } }}
      />
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
      />
    </Stack>
  );

  if (!isAccountant) {
    return (
      <Stack tokens={{ childrenGap: 12 }} styles={{ root: { padding: 16 } }}>
        {tableContent}
      </Stack>
    );
  }

  // Accountant: two-tab view
  return (
    <Stack tokens={{ childrenGap: 12 }} styles={{ root: { padding: 16 } }}>
      <Pivot>
        <PivotItem headerText={strings.TabDocuments}>
          <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 16 } }}>
            {tableContent}
          </Stack>
        </PivotItem>
        <PivotItem headerText={strings.TabUpload}>
          <Stack styles={{ root: { marginTop: 16 } }}>
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
          </Stack>
        </PivotItem>
      </Pivot>
    </Stack>
  );
};

export default SalaryDocumentWebPartDashboard;
