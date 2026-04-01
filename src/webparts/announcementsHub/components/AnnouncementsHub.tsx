import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { PrimaryButton, IconButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';

import { IAnnouncementsHubProps, IAnnouncement } from './IAnnouncementsHubProps';
import { AnnouncementsService } from '../../../services/AnnouncementsService';

import AnnouncementGrid from './AnnouncementGrid';
import AnnouncementList from './AnnouncementList';
import AnnouncementForm from './AnnouncementForm';
import styles from './AnnouncementsHub.module.scss';

const ALL_CATEGORIES = ['All', 'General', 'HR', 'IT', 'Finance', 'Events', 'Policy', 'Urgent'];

const AnnouncementsHub: React.FC<IAnnouncementsHubProps> = (props) => {
  const {
    title,
    layoutMode,
    itemsToDisplay,
    sortOrder,
    showImages,
    enableAnimations,
    enableCategoryColors,
    accentColor,
    spHttpClient,
    siteUrl,
    currentUserId,
    isAdmin
  } = props;

  // ─── State ──────────────────────────────────────────────────────────
  const [announcements, setAnnouncements] = React.useState<IAnnouncement[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string>('');
  const [selectedCategory, setSelectedCategory] = React.useState<string>('All');
  const [currentLayout, setCurrentLayout] = React.useState<'grid' | 'list'>(layoutMode);
  const [formOpen, setFormOpen] = React.useState<boolean>(false);
  const [selectedAnnouncement, setSelectedAnnouncement] = React.useState<IAnnouncement | null>(null);

  // ─── Service ────────────────────────────────────────────────────────
  const serviceRef = React.useRef<AnnouncementsService | null>(null);
  if (!serviceRef.current) {
    serviceRef.current = new AnnouncementsService(spHttpClient, siteUrl);
  }

  // ─── Sync layout mode from property pane ────────────────────────────
  React.useEffect(() => {
    setCurrentLayout(layoutMode);
  }, [layoutMode]);

  // ─── Data Fetching ──────────────────────────────────────────────────
  const fetchData = React.useCallback(async (): Promise<void> => {
    if (!serviceRef.current) return;
    setLoading(true);
    setError('');
    try {
      const data = await serviceRef.current.getAnnouncements(
        itemsToDisplay,
        sortOrder,
        currentUserId
      );
      setAnnouncements(data);
    } catch (err) {
      console.error('AnnouncementsHub: fetch error', err);
      setError('Failed to load announcements. Please try again.');
    } finally {
      setLoading(false);
    }
  }, [itemsToDisplay, sortOrder, currentUserId]);

  React.useEffect(() => {
    fetchData().catch(() => { /* handled */ });
  }, [fetchData]);

  // ─── Handlers ───────────────────────────────────────────────────────
  const handleDismiss = React.useCallback(async (id: number): Promise<void> => {
    if (!serviceRef.current) return;
    try {
      await serviceRef.current.dismissAnnouncement(id, currentUserId);
      setAnnouncements(prev => prev.map(a => a.Id === id ? { ...a, IsDismissed: true } : a));
    } catch (err) {
      console.error('AnnouncementsHub: dismiss error', err);
    }
  }, [currentUserId]);

  const handleDelete = React.useCallback(async (id: number): Promise<void> => {
    if (!serviceRef.current) return;
    try {
      await serviceRef.current.deleteAnnouncement(id);
      setAnnouncements(prev => prev.filter(a => a.Id !== id));
    } catch (err) {
      console.error('AnnouncementsHub: delete error', err);
    }
  }, []);

  const handleCreate = React.useCallback(async (
    announcementTitle: string,
    description: string,
    category: string,
    imageUrl: string,
    isImportant: boolean
  ): Promise<void> => {
    if (!serviceRef.current) return;
    await serviceRef.current.createAnnouncement(announcementTitle, description, category, imageUrl, isImportant);
    await fetchData();
  }, [fetchData]);

  // ─── Filtered data ─────────────────────────────────────────────────
  const filteredAnnouncements = React.useMemo(() => {
    let items = announcements.filter(a => !a.IsDismissed);
    if (selectedCategory !== 'All') {
      items = items.filter(a => a.Category === selectedCategory);
    }
    return items;
  }, [announcements, selectedCategory]);

  // ─── Apply accent color CSS variable ────────────────────────────────
  const rootStyle = React.useMemo(
    () => ({ '--accent': accentColor } as React.CSSProperties),
    [accentColor]
  );

  // ─── Shimmer Skeleton ──────────────────────────────────────────────
  const renderShimmer = (): React.ReactElement => (
    <div className={styles.shimmerGrid}>
      {[1, 2, 3].map(i => (
        <div key={i} className={styles.shimmerCard}>
          <div className={styles.shimmerImage} />
          <div className={styles.shimmerLine} />
          <div className={styles.shimmerLine} />
          <div className={styles.shimmerLine} />
        </div>
      ))}
    </div>
  );

  // ─── Empty State ───────────────────────────────────────────────────
  const renderEmptyState = (): React.ReactElement => (
    <div className={styles.emptyState} role="status">
      <div className={styles.emptyIcon}>
        <Icon iconName="Megaphone" />
      </div>
      <h3>No announcements available</h3>
      <p>
        {selectedCategory !== 'All'
          ? `There are no announcements in the "${selectedCategory}" category.`
          : 'Check back later for new organizational announcements.'}
      </p>
    </div>
  );

  // ─── Render ─────────────────────────────────────────────────────────
  return (
    <section className={styles.announcementsHub} style={rootStyle} aria-label={title}>
      {/* Header */}
      <div className={styles.header}>
        <div className={styles.headerLeft}>
          <h2 className={styles.title}>{title}</h2>
          {isAdmin && (
            <PrimaryButton
              text="Add Announcement"
              iconProps={{ iconName: 'Add' }}
              onClick={() => setFormOpen(true)}
            />
          )}
        </div>

        <div className={styles.headerActions}>
          {/* View Toggle */}
          <IconButton
            iconProps={{ iconName: 'GridViewMedium' }}
            title="Grid view"
            ariaLabel="Grid view"
            onClick={() => setCurrentLayout('grid')}
            checked={currentLayout === 'grid'}
          />
          <IconButton
            iconProps={{ iconName: 'List' }}
            title="List view"
            ariaLabel="List view"
            onClick={() => setCurrentLayout('list')}
            checked={currentLayout === 'list'}
          />
        </div>
      </div>

      {/* Category Filter Pills */}
      <div className={styles.filterBar} role="tablist" aria-label="Filter by category">
        {ALL_CATEGORIES.map(cat => (
          <button
            key={cat}
            className={selectedCategory === cat ? styles.filterPillActive : styles.filterPill}
            onClick={() => setSelectedCategory(cat)}
            role="tab"
            aria-selected={selectedCategory === cat}
            aria-label={`Filter: ${cat}`}
          >
            {cat}
          </button>
        ))}
      </div>

      {/* Content */}
      {loading ? (
        renderShimmer()
      ) : error ? (
        <div className={styles.emptyState} role="alert">
          <div className={styles.emptyIcon}>
            <Icon iconName="ErrorBadge" />
          </div>
          <h3>{error}</h3>
          <PrimaryButton
            text="Retry"
            iconProps={{ iconName: 'Refresh' }}
            onClick={() => { fetchData().catch(() => { /* handled */ }); }}
            styles={{
              root: { borderRadius: 8, height: 36, marginTop: 8 },
              label: { fontWeight: 600 }
            }}
          />
        </div>
      ) : filteredAnnouncements.length === 0 ? (
        renderEmptyState()
      ) : currentLayout === 'grid' ? (
        <AnnouncementGrid
          announcements={filteredAnnouncements}
          showImages={showImages}
          enableAnimations={enableAnimations}
          enableCategoryColors={enableCategoryColors}
          accentColor={accentColor}
          isAdmin={isAdmin}
          onClick={setSelectedAnnouncement}
          onDismiss={handleDismiss}
          onDelete={handleDelete}
        />
      ) : (
        <AnnouncementList
          announcements={filteredAnnouncements}
          enableAnimations={enableAnimations}
          enableCategoryColors={enableCategoryColors}
          isAdmin={isAdmin}
          onClick={setSelectedAnnouncement}
          onDismiss={handleDismiss}
          onDelete={handleDelete}
        />
      )}

      {/* Admin Form Panel */}
      {isAdmin && (
        <AnnouncementForm
          isOpen={formOpen}
          onDismiss={() => setFormOpen(false)}
          onSave={handleCreate}
        />
      )}

      {/* Detail Popover Panel */}
      <Panel
        isOpen={!!selectedAnnouncement}
        onDismiss={() => setSelectedAnnouncement(null)}
        type={PanelType.medium}
        closeButtonAriaLabel="Close"
        headerText={selectedAnnouncement?.Title}
      >
        {selectedAnnouncement && (
          <div className={styles.detailPanelContent} style={{ paddingTop: 16 }}>
            {selectedAnnouncement.IsImportant && (
              <div className={styles.importantBadge} style={{ position: 'relative', display: 'inline-flex', marginBottom: 16, top: 0, right: 0 }}>
                <Icon iconName="Warning" style={{ fontSize: 11 }} />
                Important
              </div>
            )}
            
            <div className={styles.cardMeta} style={{ marginBottom: 20 }}>
              <span className={styles.categoryBadge} style={{ background: '#f3f2f1', color: '#323130' }}>
                {selectedAnnouncement.Category}
              </span>
              <span className={styles.cardDate}>
                <Icon iconName="Calendar" style={{ fontSize: 11, marginRight: 4, opacity: 0.7 }} />
                {new Date(selectedAnnouncement.Created).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}
              </span>
            </div>

            {selectedAnnouncement.ImageUrl && (
              <img
                src={selectedAnnouncement.ImageUrl}
                alt={selectedAnnouncement.Title}
                style={{ width: '100%', borderRadius: 8, marginBottom: 20 }}
              />
            )}
            
            <p style={{ fontSize: 15, lineHeight: 1.6, color: '#323130', whiteSpace: 'pre-wrap', margin: 0 }}>
              {selectedAnnouncement.Description}
            </p>
          </div>
        )}
      </Panel>
    </section>
  );
};

export default AnnouncementsHub;
