import * as React from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  ConstrainMode
} from '@fluentui/react/lib/DetailsList';
import { Link } from '@fluentui/react/lib/Link';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Icon } from '@fluentui/react/lib/Icon';

import * as strings from 'ItSupportTicketsWebPartStrings';
import styles from './ItSupportTickets.module.scss';
import { ITicket } from '../../../services/ITopTicketService';

// ── Helpers ───────────────────────────────────────────────────────────────────

function formatDate(iso: string): string {
  if (!iso) return '';
  try {
    return new Date(iso).toLocaleDateString(undefined, {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    });
  } catch {
    return iso;
  }
}

function getStatusClass(status: string): string {
  switch (status.toLowerCase()) {
    case 'new':       return styles.statusNew;
    case 'assigned':  return styles.statusAssigned;
    case 'resolved':  return styles.statusResolved;
    case 'closed':    return styles.statusClosed;
    case 'pending':   return styles.statusPending;
    default:          return styles.statusNew;
  }
}

function getStatusLabel(status: string): string {
  switch (status.toLowerCase()) {
    case 'new':       return strings.StatusNew;
    case 'assigned':  return strings.StatusAssigned;
    case 'resolved':  return strings.StatusResolved;
    case 'closed':    return strings.StatusClosed;
    case 'pending':   return strings.StatusPending;
    default:          return status;
  }
}

function getUrgencyClass(urgency: string): string {
  switch ((urgency || '').toLowerCase()) {
    case '1': case 'critical': return styles.urgencyCritical;
    case '2': case 'high':     return styles.urgencyHigh;
    case '3': case 'medium':   return styles.urgencyMedium;
    case '4': case 'low':      return styles.urgencyLow;
    default:                   return styles.urgencyLow;
  }
}

function getUrgencyLabel(ticket: ITicket): string {
  if (ticket.urgencyLabel) return ticket.urgencyLabel;
  switch (String(ticket.urgency)) {
    case '1': return strings.UrgencyCritical;
    case '2': return strings.UrgencyHigh;
    case '3': return strings.UrgencyMedium;
    case '4': return strings.UrgencyLow;
    default:  return ticket.urgency || '';
  }
}

// ── Props ─────────────────────────────────────────────────────────────────────

export interface ITicketTableProps {
  items: ITicket[];
  loading: boolean;
  hasMore?: boolean;
  onLoadMore?: () => void;
  emptyMessage?: string;
}

// ── Component ─────────────────────────────────────────────────────────────────

export const TicketTable: React.FC<ITicketTableProps> = ({
  items,
  loading,
  hasMore,
  onLoadMore,
  emptyMessage
}) => {
  const columns = React.useMemo<IColumn[]>(() => [
    {
      key: 'ref',
      name: strings.ColumnRef,
      fieldName: 'ref',
      minWidth: 80,
      maxWidth: 100,
      isResizable: true,
      onRender: (item: ITicket) => (
        <Link
          href={item.deepLink}
          target="_blank"
          rel="noreferrer noopener"
          title={`Open ticket ${item.ref} in iTop`}
          styles={{ root: { fontWeight: 600, whiteSpace: 'nowrap' } }}
        >
          {item.ref}
        </Link>
      )
    },
    {
      key: 'title',
      name: strings.ColumnTitle,
      fieldName: 'title',
      minWidth: 160,
      maxWidth: 360,
      isResizable: true,
      isMultiline: false,
      onRender: (item: ITicket) => (
        <Text
          styles={{
            root: {
              overflow: 'hidden',
              textOverflow: 'ellipsis',
              whiteSpace: 'nowrap',
              display: 'block',
              maxWidth: '100%'
            }
          }}
          title={item.title}
        >
          {item.title}
        </Text>
      )
    },
    {
      key: 'status',
      name: strings.ColumnStatus,
      fieldName: 'status',
      minWidth: 90,
      maxWidth: 120,
      isResizable: true,
      onRender: (item: ITicket) => (
        <span
          className={`${styles.statusBadge} ${getStatusClass(item.status)}`}
        >
          {getStatusLabel(item.status)}
        </span>
      )
    },
    {
      key: 'urgency',
      name: strings.ColumnUrgency,
      fieldName: 'urgency',
      minWidth: 80,
      maxWidth: 110,
      isResizable: true,
      onRender: (item: ITicket) => (
        <span
          className={`${styles.urgencyBadge} ${getUrgencyClass(item.urgency)}`}
        >
          {getUrgencyLabel(item)}
        </span>
      )
    },
    {
      key: 'service',
      name: strings.ColumnService,
      fieldName: 'service',
      minWidth: 120,
      maxWidth: 200,
      isResizable: true,
      onRender: (item: ITicket) => (
        <Text variant="small" title={item.service}>
          {item.service}
        </Text>
      )
    },
    {
      key: 'lastUpdate',
      name: strings.ColumnLastUpdate,
      fieldName: 'lastUpdate',
      minWidth: 100,
      maxWidth: 130,
      isResizable: true,
      onRender: (item: ITicket) => (
        <Text variant="small">{formatDate(item.lastUpdate)}</Text>
      )
    }
  ], []);

  // ── Empty state ─────────────────────────────────────────────────────────────
  if (!loading && items.length === 0) {
    return (
      <Stack horizontalAlign="center" className={styles.emptyState}>
        <Icon iconName="Repair" className={styles.emptyStateIcon} />
        <Text variant="large">
          {emptyMessage ?? strings.NoTicketsMessage}
        </Text>
      </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <div className={styles.tableScrollWrapper}>
        <DetailsList
          items={items}
          columns={columns}
          layoutMode={DetailsListLayoutMode.fixedColumns}
          constrainMode={ConstrainMode.unconstrained}
          selectionMode={SelectionMode.none}
          isHeaderVisible={true}
          compact={false}
        />
      </div>

      {loading && (
        <Stack horizontalAlign="center" styles={{ root: { padding: '8px 0' } }}>
          <Spinner size={SpinnerSize.small} />
        </Stack>
      )}

      {!loading && hasMore && onLoadMore && (
        <Stack horizontalAlign="center" className={styles.loadMoreWrapper}>
          <DefaultButton text={strings.LoadMoreLabel} onClick={onLoadMore} />
        </Stack>
      )}
    </Stack>
  );
};

export default TicketTable;
