import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import { IAnnouncement } from './IAnnouncementsHubProps';
import styles from './AnnouncementsHub.module.scss';

/** Category → color mapping (same as AnnouncementCard) */
const CATEGORY_COLORS: Record<string, { bg: string; text: string; glow: string }> = {
  General:  { bg: '#e8e6e4', text: '#3b3a39', glow: 'rgba(59,58,57,0.10)' },
  HR:       { bg: '#f0e0f5', text: '#8b2fa2', glow: 'rgba(139,47,162,0.12)' },
  IT:       { bg: '#d4e8ff', text: '#0050a0', glow: 'rgba(0,80,160,0.12)' },
  Finance:  { bg: '#d4f0da', text: '#107c41', glow: 'rgba(16,124,65,0.12)' },
  Events:   { bg: '#fff4d1', text: '#9a6700', glow: 'rgba(154,103,0,0.12)' },
  Policy:   { bg: '#fde8ea', text: '#c42b30', glow: 'rgba(196,43,48,0.12)' },
  Urgent:   { bg: '#d13438', text: '#ffffff', glow: 'rgba(209,52,56,0.25)' }
};

export interface IAnnouncementListProps {
  announcements: IAnnouncement[];
  enableAnimations: boolean;
  enableCategoryColors: boolean;
  isAdmin: boolean;
  onClick?: (announcement: IAnnouncement) => void;
  onDismiss?: (id: number) => void;
  onDelete?: (id: number) => void;
}

const AnnouncementList: React.FC<IAnnouncementListProps> = React.memo((props) => {
  const {
    announcements,
    enableAnimations,
    enableCategoryColors,
    isAdmin,
    onClick,
    onDismiss,
    onDelete
  } = props;

  const defaultCat = { bg: '#e8e6e4', text: '#3b3a39', glow: 'rgba(59,58,57,0.10)' };

  return (
    <div className={styles.list} role="list" aria-label="Announcements list">
      {announcements.map((a) => {
        const dateStr = new Date(a.Created).toLocaleDateString('en-US', {
          month: 'short', day: 'numeric', year: 'numeric'
        });
        const catColor = enableCategoryColors
          ? (CATEGORY_COLORS[a.Category] || defaultCat)
          : defaultCat;
        const rowClass = enableAnimations ? styles.listRowAnimated : styles.listRow;

        return (
          <div
            key={a.Id}
            className={rowClass}
            role="button"
            aria-label={`View Announcement: ${a.Title}`}
            tabIndex={0}
            onClick={() => { if (onClick) onClick(a); }}
            onKeyPress={(e) => {
              if (e.key === 'Enter' || e.key === ' ') {
                if (onClick) onClick(a);
              }
            }}
          >
            {/* Important dot */}
            {a.IsImportant && (
              <div
                className={styles.listImportantDot}
                title="Important announcement"
                aria-label="Important"
              />
            )}

            {/* Category badge */}
            <div className={styles.listBadge}>
              <span
                className={styles.categoryBadge}
                style={{
                  backgroundColor: catColor.bg,
                  color: catColor.text,
                  boxShadow: `0 1px 6px ${catColor.glow}`
                }}
              >
                {a.Category}
              </span>
            </div>

            {/* Content */}
            <div className={styles.listContent}>
              <div className={styles.listTitle}>{a.Title}</div>
              {a.Description && (
                <div className={styles.listDescription}>{a.Description}</div>
              )}
            </div>

            {/* Date */}
            <span className={styles.listDate}>
              <Icon iconName="Calendar" style={{ marginRight: 4, fontSize: 11, opacity: 0.7 }} />
              {dateStr}
            </span>

            {/* Actions */}
            <div className={styles.listActions}>
              {onDismiss && (
                <IconButton
                  iconProps={{ iconName: 'Cancel' }}
                  title="Dismiss"
                  ariaLabel="Dismiss announcement"
                  styles={{
                    root: { width: 30, height: 30, borderRadius: 8 },
                    icon: { fontSize: 12, color: '#8a8886' },
                    rootHovered: { background: '#f3f2f1' }
                  }}
                  onClick={(e) => {
                    e.stopPropagation();
                    if (onDismiss) onDismiss(a.Id);
                  }}
                />
              )}
              {isAdmin && onDelete && (
                <IconButton
                  iconProps={{ iconName: 'Delete' }}
                  title="Delete"
                  ariaLabel="Delete announcement"
                  styles={{
                    root: { width: 30, height: 30, borderRadius: 8 },
                    icon: { fontSize: 12, color: '#d13438' },
                    rootHovered: { background: '#fde8ea', color: '#a80000' }
                  }}
                  onClick={(e) => {
                    e.stopPropagation();
                    if (onDelete) onDelete(a.Id);
                  }}
                />
              )}
            </div>
          </div>
        );
      })}
    </div>
  );
});

AnnouncementList.displayName = 'AnnouncementList';

export default AnnouncementList;
