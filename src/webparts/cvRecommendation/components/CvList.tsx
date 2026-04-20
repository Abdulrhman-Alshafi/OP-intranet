import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton, DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Stack } from '@fluentui/react/lib/Stack';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';

import {
  ICvRecommendation,
  CvRecommendationService
} from '../../../services/CvRecommendationService';
import styles from './CvRecommendation.module.scss';

// ─── Props ────────────────────────────────────────────────────────────────────
export interface ICvListProps {
  items: ICvRecommendation[];
  isHr: boolean;
  currentUserId: number;
  itemsPerPage: number;
  service: CvRecommendationService;
  onRefresh: () => void;
  accentColor: string;
}

// ─── Constants ────────────────────────────────────────────────────────────────
const STATUS_OPTIONS: IDropdownOption[] = [
  { key: 'All', text: 'All Statuses' },
  { key: 'Submitted', text: 'Submitted' },
  { key: 'Under Review', text: 'Under Review' },
  { key: 'Accepted', text: 'Accepted' },
  { key: 'Rejected', text: 'Rejected' }
];

type SortField = 'Created' | 'Title' | 'Position' | 'Status';
type SortDir = 'asc' | 'desc';

// ─── Helpers ──────────────────────────────────────────────────────────────────
function getStatusClass(status: string): string {
  switch (status) {
    case 'Submitted':   return styles.statusSubmitted;
    case 'Under Review': return styles.statusUnderReview;
    case 'Accepted':    return styles.statusAccepted;
    case 'Rejected':    return styles.statusRejected;
    default:            return '';
  }
}

function formatDate(iso: string): string {
  if (!iso) return '';
  try {
    return new Date(iso).toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
  } catch { return iso; }
}

// ─── CvList Component ─────────────────────────────────────────────────────────
const CvList: React.FC<ICvListProps> = ({
  items, isHr, currentUserId, itemsPerPage, service, onRefresh
}) => {
  // ─── Filter / Search ────────────────────────────────────────────
  const [search, setSearch] = React.useState('');
  const [statusFilter, setStatusFilter] = React.useState<string>('All');

  // ─── Sort ────────────────────────────────────────────────────────
  const [sortField, setSortField] = React.useState<SortField>('Created');
  const [sortDir, setSortDir] = React.useState<SortDir>('desc');

  // ─── Pagination ──────────────────────────────────────────────────
  const [page, setPage] = React.useState(1);

  // ─── Detail Panel ────────────────────────────────────────────────
  const [selected, setSelected] = React.useState<ICvRecommendation | null>(null);
  const [detailOpen, setDetailOpen] = React.useState(false);

  // ─── Status Update (HR) ──────────────────────────────────────────
  const [updatingStatus, setUpdatingStatus] = React.useState(false);
  const [statusError, setStatusError] = React.useState('');

  // ─── Delete Confirm ──────────────────────────────────────────────
  const [deleteTarget, setDeleteTarget] = React.useState<ICvRecommendation | null>(null);
  const [deleting, setDeleting] = React.useState(false);

  // ─── Computed / Filtered ─────────────────────────────────────────
  const filtered = React.useMemo(() => {
    let result = [...items];

    if (search.trim()) {
      const q = search.trim().toLowerCase();
      result = result.filter(i =>
        i.Title.toLowerCase().includes(q) ||
        i.CandidateEmail.toLowerCase().includes(q) ||
        i.Position.toLowerCase().includes(q) ||
        i.SubmittedByName.toLowerCase().includes(q)
      );
    }

    if (statusFilter !== 'All') {
      result = result.filter(i => i.Status === statusFilter);
    }

    result.sort((a, b) => {
      let va: string;
      let vb: string;
      if (sortField === 'Created') {
        va = a.Created; vb = b.Created;
      } else {
        va = (a[sortField] ?? '').toString();
        vb = (b[sortField] ?? '').toString();
      }
      const cmp = va.localeCompare(vb);
      return sortDir === 'asc' ? cmp : -cmp;
    });

    return result;
  }, [items, search, statusFilter, sortField, sortDir]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / itemsPerPage));
  const currentPage = Math.min(page, totalPages);
  const pageItems = filtered.slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage);

  // Reset to page 1 when filters change
  React.useEffect(() => { setPage(1); }, [search, statusFilter]);

  // ─── Handlers ────────────────────────────────────────────────────
  const handleSort = React.useCallback((field: SortField): void => {
    setSortField(prev => {
      if (prev === field) { setSortDir(d => d === 'asc' ? 'desc' : 'asc'); return prev; }
      setSortDir('asc');
      return field;
    });
  }, []);

  const openDetail = React.useCallback((item: ICvRecommendation): void => {
    setSelected(item);
    setStatusError('');
    setDetailOpen(true);
  }, []);

  const closeDetail = React.useCallback((): void => {
    setDetailOpen(false);
    setSelected(null);
    setStatusError('');
  }, []);

  const handleStatusChange = React.useCallback(async (
    item: ICvRecommendation,
    newStatus: ICvRecommendation['Status']
  ): Promise<void> => {
    if (!isHr) return;
    setUpdatingStatus(true);
    setStatusError('');
    try {
      await service.updateStatus(item.Id, newStatus);
      setSelected(prev => prev ? { ...prev, Status: newStatus } : prev);
      onRefresh();
    } catch (e) {
      setStatusError(e instanceof Error ? e.message : 'Failed to update status.');
    } finally {
      setUpdatingStatus(false);
    }
  }, [isHr, service, onRefresh]);

  const handleDeleteConfirm = React.useCallback(async (): Promise<void> => {
    if (!deleteTarget) return;
    setDeleting(true);
    try {
      await service.deleteCvRecommendation(deleteTarget.Id);
      setDeleteTarget(null);
      onRefresh();
      if (selected?.Id === deleteTarget.Id) closeDetail();
    } catch (e) {
      console.error('CvList: delete error', e);
    } finally {
      setDeleting(false);
    }
  }, [deleteTarget, service, onRefresh, selected, closeDetail]);

  // ─── Sort indicator ──────────────────────────────────────────────
  const sortIcon = (field: SortField): string => {
    if (sortField !== field) return 'Sort';
    return sortDir === 'asc' ? 'SortUp' : 'SortDown';
  };

  // ─── Render Status Pills ─────────────────────────────────────────
  const renderStatusActions = (item: ICvRecommendation): JSX.Element => {
    const statuses: ICvRecommendation['Status'][] = ['Submitted', 'Under Review', 'Accepted', 'Rejected'];
    return (
      <Stack horizontal tokens={{ childrenGap: 6 }} wrap>
        {statusError && (
          <MessageBar messageBarType={MessageBarType.error} style={{ marginBottom: 8 }}>
            {statusError}
          </MessageBar>
        )}
        {statuses.map(s => (
          <span
            key={s}
            className={`${styles.statusBadge} ${getStatusClass(s)}`}
            style={{
              cursor: item.Status === s || updatingStatus ? 'default' : 'pointer',
              opacity: item.Status === s ? 1 : 0.55,
              outline: item.Status === s ? '2px solid currentColor' : 'none',
              outlineOffset: 1
            }}
            onClick={() => { if (item.Status !== s && !updatingStatus) handleStatusChange(item, s).catch(() => undefined); }}
            role="button"
            tabIndex={0}
            title={`Set status to ${s}`}
            onKeyDown={(e) => { if (e.key === 'Enter') handleStatusChange(item, s).catch(() => undefined); }}
          >
            {s}{updatingStatus && item.Status !== s ? '' : ''}
          </span>
        ))}
      </Stack>
    );
  };

  // ─── Render ──────────────────────────────────────────────────────
  return (
    <>
      {/* Toolbar */}
      <div className={styles.toolbar}>
        <div className={styles.toolbarLeft}>
          <SearchBox
            placeholder="Search candidate, position…"
            className={styles.searchBox}
            value={search}
            onChange={(_, v) => setSearch(v ?? '')}
            onClear={() => setSearch('')}
          />
          <Dropdown
            options={STATUS_OPTIONS}
            selectedKey={statusFilter}
            onChange={(_, o) => setStatusFilter(o?.key as string ?? 'All')}
            styles={{ root: { minWidth: 160 } }}
          />
        </div>
        <div className={styles.toolbarRight}>
          <span style={{ fontSize: 12, color: 'var(--neutralSecondary, #605e5c)' }}>
            {filtered.length} result{filtered.length !== 1 ? 's' : ''}
          </span>
        </div>
      </div>

      {/* Table */}
      {pageItems.length === 0 ? (
        <div className={styles.stateBox}>
          <Icon iconName="SearchIssue" className={styles.stateIcon} />
          <div className={styles.stateTitle}>No results found</div>
          <div className={styles.stateDesc}>
            {search || statusFilter !== 'All'
              ? 'Try adjusting your search or filter.'
              : 'No CV submissions yet.'}
          </div>
        </div>
      ) : (
        <div className={styles.tableWrapper}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th onClick={() => handleSort('Title')}>
                  Candidate&nbsp;<Icon iconName={sortIcon('Title')} style={{ fontSize: 11 }} />
                </th>
                <th>Email</th>
                <th onClick={() => handleSort('Position')}>
                  Position&nbsp;<Icon iconName={sortIcon('Position')} style={{ fontSize: 11 }} />
                </th>
                {isHr && <th>Submitted By</th>}
                <th onClick={() => handleSort('Created')}>
                  Date&nbsp;<Icon iconName={sortIcon('Created')} style={{ fontSize: 11 }} />
                </th>
                <th onClick={() => handleSort('Status')}>
                  Status&nbsp;<Icon iconName={sortIcon('Status')} style={{ fontSize: 11 }} />
                </th>
                <th>Actions</th>
              </tr>
            </thead>
            <tbody>
              {pageItems.map(item => {
                const ownItem = item.SubmittedById === currentUserId;
                const canDelete = isHr || ownItem;

                return (
                  <tr key={item.Id}>
                    <td>
                      <span
                        style={{ cursor: 'pointer', fontWeight: 600, color: 'var(--themePrimary, #0078d4)' }}
                        onClick={() => openDetail(item)}
                        role="button"
                        tabIndex={0}
                        onKeyDown={(e) => { if (e.key === 'Enter') openDetail(item); }}
                      >
                        {item.Title}
                      </span>
                    </td>
                    <td>{item.CandidateEmail}</td>
                    <td>{item.Position}</td>
                    {isHr && <td>{item.SubmittedByName}</td>}
                    <td>{formatDate(item.Created)}</td>
                    <td>
                      <span className={`${styles.statusBadge} ${getStatusClass(item.Status)}`}>
                        {item.Status}
                      </span>
                    </td>
                    <td>
                      <div className={styles.actionsCell}>
                        <IconButton
                          iconProps={{ iconName: 'View' }}
                          title="View details"
                          onClick={() => openDetail(item)}
                        />
                        {item.CVFileUrl && (
                          <IconButton
                            iconProps={{ iconName: 'Download' }}
                            title="Download CV"
                            href={item.CVFileUrl}
                            target="_blank"
                            rel="noopener noreferrer"
                          />
                        )}
                        {canDelete && (
                          <IconButton
                            iconProps={{ iconName: 'Delete' }}
                            title="Delete"
                            onClick={() => setDeleteTarget(item)}
                            styles={{ root: { color: '#a4262c' } }}
                          />
                        )}
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {/* Pagination */}
      {totalPages > 1 && (
        <div className={styles.pagination}>
          <span className={styles.pageInfo}>
            Page {currentPage} of {totalPages} &bull; {filtered.length} items
          </span>
          <div className={styles.pageButtons}>
            <DefaultButton
              iconProps={{ iconName: 'ChevronLeft' }}
              disabled={currentPage === 1}
              onClick={() => setPage(p => p - 1)}
              text="Prev"
            />
            <DefaultButton
              iconProps={{ iconName: 'ChevronRight' }}
              disabled={currentPage === totalPages}
              onClick={() => setPage(p => p + 1)}
              text="Next"
            />
          </div>
        </div>
      )}

      {/* Detail Panel */}
      {selected && (
        <Panel
          isOpen={detailOpen}
          onDismiss={closeDetail}
          type={PanelType.medium}
          headerText="CV Recommendation Details"
          closeButtonAriaLabel="Close"
        >
          <Stack tokens={{ childrenGap: 12 }} style={{ paddingTop: 16 }}>

            {/* Status Badge */}
            <div style={{ marginBottom: 8 }}>
              <span className={`${styles.statusBadge} ${getStatusClass(selected.Status)}`} style={{ fontSize: 13 }}>
                {selected.Status}
              </span>
            </div>

            <div className={styles.detailRow}>
              <span className={styles.detailLabel}>Candidate</span>
              <span className={styles.detailValue}>{selected.Title}</span>
            </div>
            <div className={styles.detailRow}>
              <span className={styles.detailLabel}>Email</span>
              <span className={styles.detailValue}>
                <a href={`mailto:${selected.CandidateEmail}`}>{selected.CandidateEmail}</a>
              </span>
            </div>
            {selected.PhoneNumber && (
              <div className={styles.detailRow}>
                <span className={styles.detailLabel}>Phone</span>
                <span className={styles.detailValue}>{selected.PhoneNumber}</span>
              </div>
            )}
            <div className={styles.detailRow}>
              <span className={styles.detailLabel}>Position</span>
              <span className={styles.detailValue}>{selected.Position}</span>
            </div>
            {isHr && (
              <div className={styles.detailRow}>
                <span className={styles.detailLabel}>Submitted By</span>
                <span className={styles.detailValue}>{selected.SubmittedByName}</span>
              </div>
            )}
            <div className={styles.detailRow}>
              <span className={styles.detailLabel}>Date</span>
              <span className={styles.detailValue}>{formatDate(selected.Created)}</span>
            </div>
            {selected.Notes && (
              <div className={styles.detailRow}>
                <span className={styles.detailLabel}>Notes</span>
                <span className={styles.detailValue} style={{ whiteSpace: 'pre-wrap' }}>{selected.Notes}</span>
              </div>
            )}
            {selected.CVFileUrl && (
              <div className={styles.detailRow}>
                <span className={styles.detailLabel}>CV File</span>
                <span className={styles.detailValue}>
                  <a href={selected.CVFileUrl} target="_blank" rel="noopener noreferrer">
                    <Icon iconName="Attach" style={{ marginRight: 4 }} />
                    {selected.CVFileName || 'Download'}
                  </a>
                </span>
              </div>
            )}

            {/* HR Status Actions */}
            {isHr && (
              <div className={styles.formSection} style={{ marginTop: 12 }}>
                <p className={styles.formSectionTitle}>Update Status</p>
                {renderStatusActions(selected)}
              </div>
            )}

          </Stack>
        </Panel>
      )}

      {/* Delete Confirmation Dialog */}
      <Dialog
        hidden={!deleteTarget}
        onDismiss={() => setDeleteTarget(null)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Delete CV Recommendation',
          subText: deleteTarget
            ? `Are you sure you want to delete the CV recommendation for "${deleteTarget.Title}"? This action cannot be undone.`
            : ''
        }}
        modalProps={{ isBlocking: true }}
      >
        <DialogFooter>
          <PrimaryButton
            text={deleting ? 'Deleting…' : 'Delete'}
            onClick={handleDeleteConfirm}
            disabled={deleting}
            styles={{ root: { backgroundColor: '#a4262c', borderColor: '#a4262c' } }}
          />
          <DefaultButton text="Cancel" onClick={() => setDeleteTarget(null)} disabled={deleting} />
        </DialogFooter>
      </Dialog>
    </>
  );
};

export default CvList;
