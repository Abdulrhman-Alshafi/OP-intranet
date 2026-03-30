import * as React from 'react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';
import { SPHttpClient } from '@microsoft/sp-http';

import * as strings from 'ItSupportTicketsWebPartStrings';
import { IItSupportTicketsProps } from './IItSupportTicketsProps';
import { TicketTable } from './TicketTable';
import { NewTicketPanel } from './NewTicketPanel';
import { ITicket } from '../../../services/ITopTicketService';
import styles from './ItSupportTickets.module.scss';

const PAGE_SIZE = 50;

// ── SP group check helper ─────────────────────────────────────────────────────

async function isUserInGroup(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  groupName: string
): Promise<boolean> {
  try {
    const url = `${siteUrl}/_api/web/currentuser/groups?$select=LoginName,Title`;
    const response = await spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) return false;
    const data: { value: { Title: string }[] } = await response.json();
    const groups = data?.value ?? [];
    return groups.some(
      (g) => g.Title.toLowerCase() === groupName.toLowerCase()
    );
  } catch {
    return false;
  }
}

// ── Status filter options ─────────────────────────────────────────────────────

const STATUS_OPTIONS: IDropdownOption[] = [
  { key: '', text: strings.FilterAllStatuses },
  { key: 'new', text: strings.StatusNew },
  { key: 'assigned', text: strings.StatusAssigned },
  { key: 'pending', text: strings.StatusPending },
  { key: 'resolved', text: strings.StatusResolved },
  { key: 'closed', text: strings.StatusClosed }
];

// ── Component ─────────────────────────────────────────────────────────────────

export const TicketDashboard: React.FC<IItSupportTicketsProps> = (props) => {
  const {
    service,
    isConfigured,
    itAdminGroupName,
    currentUserEmail,
    siteUrl,
    spHttpClient
  } = props;

  // ── Auth state ────────────────────────────────────────────────────────────
  const [isAdmin, setIsAdmin] = useState(false);
  const [authLoading, setAuthLoading] = useState(true);
  const [authError, setAuthError] = useState<string | undefined>(undefined);

  // ── My Tickets state ──────────────────────────────────────────────────────
  const [myTickets, setMyTickets] = useState<ITicket[]>([]);
  const [myLoading, setMyLoading] = useState(false);
  const [myError, setMyError] = useState<string | undefined>(undefined);
  const [mySearch, setMySearch] = useState('');
  const [myStatusFilter, setMyStatusFilter] = useState('');

  // ── All Tickets state (admin) ─────────────────────────────────────────────
  const [allTickets, setAllTickets] = useState<ITicket[]>([]);
  const [allHasMore, setAllHasMore] = useState(false);
  const [allPage, setAllPage] = useState(1);
  const [allLoading, setAllLoading] = useState(false);
  const [allError, setAllError] = useState<string | undefined>(undefined);
  const [allSearch, setAllSearch] = useState('');
  const [allStatusFilter, setAllStatusFilter] = useState('');

  const mountedRef = useRef(true);

  // ── Auth check on mount ───────────────────────────────────────────────────
  useEffect(() => {
    mountedRef.current = true;
    let cancelled = false;

    setAuthLoading(true);
    setAuthError(undefined);

    (async () => {
      try {
        const admin = await isUserInGroup(spHttpClient, siteUrl, itAdminGroupName);
        if (!cancelled) setIsAdmin(admin);
      } catch {
        if (!cancelled) setAuthError(strings.ErrorAuthCheck);
      } finally {
        if (!cancelled) setAuthLoading(false);
      }
    })().catch(() => undefined);

    return () => {
      cancelled = true;
      mountedRef.current = false;
    };
  }, [spHttpClient, siteUrl, itAdminGroupName]);

  // ── Load My Tickets once auth is resolved ─────────────────────────────────
  useEffect(() => {
    if (authLoading) return;
    let cancelled = false;

    setMyLoading(true);
    setMyError(undefined);

    (async () => {
      try {
        const items = await service.getMyTickets(currentUserEmail);
        if (!cancelled) setMyTickets(items);
      } catch (err) {
        if (!cancelled) {
          setMyError(strings.ErrorLoadTickets);
          console.error('[TicketDashboard] getMyTickets:', err);
        }
      } finally {
        if (!cancelled) setMyLoading(false);
      }
    })().catch(() => undefined);

    return () => { cancelled = true; };
  }, [authLoading, service, currentUserEmail]);

  // ── Load All Tickets (admin, first page) ─────────────────────────────────
  useEffect(() => {
    if (authLoading || !isAdmin) return;
    let cancelled = false;

    setAllLoading(true);
    setAllError(undefined);
    setAllPage(1);

    (async () => {
      try {
        const result = await service.getAllTickets(1, PAGE_SIZE);
        if (!cancelled) {
          setAllTickets(result.items);
          setAllHasMore(result.hasMore);
        }
      } catch (err) {
        if (!cancelled) {
          setAllError(strings.ErrorLoadTickets);
          console.error('[TicketDashboard] getAllTickets:', err);
        }
      } finally {
        if (!cancelled) setAllLoading(false);
      }
    })().catch(() => undefined);

    return () => { cancelled = true; };
  }, [authLoading, isAdmin, service]);

  // ── Load next page (All Tickets) ──────────────────────────────────────────
  const handleLoadMoreAll = useCallback(async () => {
    const nextPage = allPage + 1;
    setAllLoading(true);
    try {
      const result = await service.getAllTickets(nextPage, PAGE_SIZE);
      if (mountedRef.current) {
        setAllTickets((prev) => [...prev, ...result.items]);
        setAllHasMore(result.hasMore);
        setAllPage(nextPage);
      }
    } catch (err) {
      if (mountedRef.current) {
        setAllError(strings.ErrorLoadTickets);
        console.error('[TicketDashboard] load more:', err);
      }
    } finally {
      if (mountedRef.current) setAllLoading(false);
    }
  }, [service, allPage]);

  // ── Refresh My Tickets after a new ticket is created ─────────────────────
  const handleNewTicketSuccess = useCallback(() => {
    setMyLoading(true);
    setMyError(undefined);
    (async () => {
      try {
        const items = await service.getMyTickets(currentUserEmail);
        if (mountedRef.current) setMyTickets(items);
      } catch {
        if (mountedRef.current) setMyError(strings.ErrorLoadTickets);
      } finally {
        if (mountedRef.current) setMyLoading(false);
      }
    })().catch(() => undefined);
  }, [service, currentUserEmail]);

  // ── Client-side filtering ─────────────────────────────────────────────────
  const filteredMyTickets = useMemo(() => {
    const q = mySearch.trim().toLowerCase();
    return myTickets.filter((t) => {
      if (myStatusFilter && t.status.toLowerCase() !== myStatusFilter) return false;
      if (!q) return true;
      return (
        t.ref.toLowerCase().includes(q) ||
        t.title.toLowerCase().includes(q) ||
        t.service.toLowerCase().includes(q)
      );
    });
  }, [myTickets, mySearch, myStatusFilter]);

  const filteredAllTickets = useMemo(() => {
    const q = allSearch.trim().toLowerCase();
    return allTickets.filter((t) => {
      if (allStatusFilter && t.status.toLowerCase() !== allStatusFilter) return false;
      if (!q) return true;
      return (
        t.ref.toLowerCase().includes(q) ||
        t.title.toLowerCase().includes(q) ||
        t.service.toLowerCase().includes(q)
      );
    });
  }, [allTickets, allSearch, allStatusFilter]);

  // ── Reusable filter bar ───────────────────────────────────────────────────
  const renderFilterBar = (
    statusFilter: string,
    onStatusChange: (v: string) => void,
    search: string,
    onSearchChange: (v: string) => void,
    placeholder: string
  ): JSX.Element => (
    <Stack
      horizontal
      wrap
      tokens={{ childrenGap: 16 }}
      styles={{ root: { rowGap: 12, alignItems: 'stretch' } }}
    >
      <Dropdown
        selectedKey={statusFilter}
        options={STATUS_OPTIONS}
        onChange={(_, o) => onStatusChange(String(o?.key ?? ''))}
        styles={{ root: { minWidth: 150, margin: 0 }, title: { height: 32, lineHeight: 30 } }}
      />
      <SearchBox
        placeholder={placeholder}
        value={search}
        onChange={(_, v) => onSearchChange(v || '')}
        styles={{ root: { flexGrow: 1, minWidth: 180, height: 32, margin: 0, padding: 0 } }}
      />
    </Stack>
  );

  // ── Render: loading auth ──────────────────────────────────────────────────
  if (authLoading) {
    return (
      <div className={styles.itSupportTickets}>
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

  // ── Render: auth error ───────────────────────────────────────────────────
  if (authError) {
    return (
      <div className={styles.itSupportTickets}>
        <div className={styles.wpHeader}>
          <div className={styles.wpTitleSection}>
            <h2 className={styles.wpTitle}>{strings.WebPartTitle}</h2>
          </div>
        </div>
        <MessageBar messageBarType={MessageBarType.error}>{authError}</MessageBar>
      </div>
    );
  }

  // ── "My Tickets" tab content ──────────────────────────────────────────────
  const myTicketsContent = (
    <Stack tokens={{ childrenGap: 8 }}>
      {renderFilterBar(
        myStatusFilter,
        setMyStatusFilter,
        mySearch,
        setMySearch,
        'Search by ref, title or service…'
      )}
      {myError && (
        <MessageBar messageBarType={MessageBarType.error}>{myError}</MessageBar>
      )}
      <TicketTable
        items={filteredMyTickets}
        loading={myLoading}
        emptyMessage={strings.NoTicketsMyMessage}
      />
    </Stack>
  );

  // ── "All Tickets" tab content (admin) ─────────────────────────────────────
  const allTicketsContent = (
    <Stack tokens={{ childrenGap: 8 }}>
      {renderFilterBar(
        allStatusFilter,
        setAllStatusFilter,
        allSearch,
        setAllSearch,
        'Search by ref, title or service…'
      )}
      {allError && (
        <MessageBar messageBarType={MessageBarType.error}>{allError}</MessageBar>
      )}
      <TicketTable
        items={filteredAllTickets}
        loading={allLoading}
        hasMore={allHasMore}
        onLoadMore={handleLoadMoreAll}
      />
    </Stack>
  );

  // ── Render: no functionBaseUrl configured ─────────────────────────────────
  if (!isConfigured) {
    return (
      <div className={styles.itSupportTickets}>
        <div className={styles.wpHeader}>
          <div className={styles.wpTitleSection}>
            <h2 className={styles.wpTitle}>{strings.WebPartTitle}</h2>
          </div>
        </div>
        <div className={styles.configWarning}>
          <Icon iconName="Warning" />
          <span>{strings.ConfigNotSet}</span>
        </div>
      </div>
    );
  }

  // ── Render: main dashboard ────────────────────────────────────────────────
  return (
    <div className={styles.itSupportTickets}>
      <div className={styles.wpHeader}>
        <div className={styles.wpTitleSection}>
          <h2 className={styles.wpTitle}>{strings.WebPartTitle}</h2>
        </div>
      </div>

      <Pivot>
        {/* My Tickets — always visible */}
        <PivotItem headerText={strings.TabMyTickets}>
          <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 16 } }}>
            {myTicketsContent}
          </Stack>
        </PivotItem>

        {/* All Tickets — IT Admin only */}
        {isAdmin && (
          <PivotItem headerText={strings.TabAllTickets}>
            <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 16 } }}>
              {allTicketsContent}
            </Stack>
          </PivotItem>
        )}

        {/* New Ticket — always visible */}
        <PivotItem headerText={strings.TabNewTicket}>
          <div style={{ marginTop: 16 }}>
            <NewTicketPanel
              service={service}
              currentUserEmail={currentUserEmail}
              onSubmitSuccess={handleNewTicketSuccess}
            />
          </div>
        </PivotItem>
      </Pivot>
    </div>
  );
};

export default TicketDashboard;
