import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import { IAnnouncement } from './IAnnouncementsHubProps';
import styles from './AnnouncementsHub.module.scss';

/** Category → color mapping */
const CATEGORY_COLORS: Record<string, { bg: string; text: string; glow: string }> = {
  General:  { bg: '#e8e6e4', text: '#3b3a39', glow: 'rgba(59,58,57,0.10)' },
  HR:       { bg: '#f0e0f5', text: '#8b2fa2', glow: 'rgba(139,47,162,0.12)' },
  IT:       { bg: '#d4e8ff', text: '#0050a0', glow: 'rgba(0,80,160,0.12)' },
  Finance:  { bg: '#d4f0da', text: '#107c41', glow: 'rgba(16,124,65,0.12)' },
  Events:   { bg: '#fff4d1', text: '#9a6700', glow: 'rgba(154,103,0,0.12)' },
  Policy:   { bg: '#fde8ea', text: '#c42b30', glow: 'rgba(196,43,48,0.12)' },
  Urgent:   { bg: '#d13438', text: '#ffffff', glow: 'rgba(209,52,56,0.25)' }
};

export interface IAnnouncementCardProps {
  announcement: IAnnouncement;
  showImages: boolean;
  enableAnimations: boolean;
  enableCategoryColors: boolean;
  accentColor: string;
  isAdmin: boolean;
  onDismiss?: (id: number) => void;
  onDelete?: (id: number) => void;
}

const AnnouncementCard: React.FC<IAnnouncementCardProps> = React.memo((props) => {
  const {
    announcement,
    showImages,
    enableAnimations,
    enableCategoryColors,
    isAdmin,
    onDismiss,
    onDelete
  } = props;

  const dateStr = React.useMemo(() => {
    const d = new Date(announcement.Created);
    return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  }, [announcement.Created]);

  const defaultCat = { bg: '#e8e6e4', text: '#3b3a39', glow: 'rgba(59,58,57,0.10)' };
  const catColor = enableCategoryColors
    ? (CATEGORY_COLORS[announcement.Category] || defaultCat)
    : defaultCat;

  const cardClass = enableAnimations ? styles.cardAnimated : styles.card;

  return (
    <article
      className={cardClass}
      role="article"
      aria-label={`Announcement: ${announcement.Title}`}
      tabIndex={0}
    >
      {/* Important badge */}
      {announcement.IsImportant && (
        <div className={styles.importantBadge}>
          <Icon iconName="Warning" style={{ fontSize: 11 }} />
          Important
        </div>
      )}

      {/* Image */}
      {showImages && (
        announcement.ImageUrl ? (
          <img
            className={styles.cardImage}
            src={announcement.ImageUrl}
            alt={announcement.Title}
            loading="lazy"
          />
        ) : (
          <div className={styles.cardImagePlaceholder}>
            <Icon iconName="Megaphone" />
          </div>
        )
      )}

      {/* Body */}
      <div className={styles.cardBody}>
        <div className={styles.cardMeta}>
          <span
            className={styles.categoryBadge}
            style={{
              backgroundColor: catColor.bg,
              color: catColor.text,
              boxShadow: `0 1px 6px ${catColor.glow}`
            }}
          >
            {announcement.Category}
          </span>
          <span className={styles.cardDate}>
            <Icon iconName="Calendar" style={{ fontSize: 11, marginRight: 4, opacity: 0.7 }} />
            {dateStr}
          </span>
        </div>

        <h3 className={styles.cardTitle}>{announcement.Title}</h3>
        <p className={styles.cardDescription}>{announcement.Description}</p>

        <div className={styles.cardFooter}>
          {onDismiss && (
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              title="Dismiss this announcement"
              ariaLabel="Dismiss announcement"
              styles={{
                root: { width: 30, height: 30, borderRadius: 8 },
                icon: { fontSize: 12, color: '#8a8886' },
                rootHovered: { background: '#f3f2f1' }
              }}
              onClick={() => onDismiss(announcement.Id)}
            />
          )}
          {isAdmin && onDelete && (
            <IconButton
              iconProps={{ iconName: 'Delete' }}
              title="Delete announcement"
              ariaLabel="Delete announcement"
              styles={{
                root: { width: 30, height: 30, borderRadius: 8 },
                icon: { fontSize: 12, color: '#d13438' },
                rootHovered: { background: '#fde8ea', color: '#a80000' }
              }}
              onClick={() => onDelete(announcement.Id)}
            />
          )}
        </div>
      </div>
    </article>
  );
});

AnnouncementCard.displayName = 'AnnouncementCard';

export default AnnouncementCard;
