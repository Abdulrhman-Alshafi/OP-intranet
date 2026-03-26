import * as React from 'react';
import importedStyles from './AnnouncementBanner.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import type { IAnnouncementBannerProps, IAnnouncement } from './IAnnouncementBannerProps';
import { escape } from '@microsoft/sp-lodash-subset';

const AnnouncementBanner = (props: IAnnouncementBannerProps): React.ReactElement => {
  const { announcements, showIcon, showDismiss, stackDirection } = props;
  const [dismissedIds, setDismissedIds] = React.useState<string[]>([]);
  const [dismissingId, setDismissingId] = React.useState<string | undefined>(undefined);

  const getIcon = (type: string): string => {
    switch (type) {
      case 'info': return 'ℹ';
      case 'warning': return '⚠';
      case 'success': return '✓';
      case 'urgent': return '!';
      default: return 'ℹ';
    }
  };

  const getBannerClass = (type: string): string => {
    switch (type) {
      case 'info': return styles.bannerInfo;
      case 'warning': return styles.bannerWarning;
      case 'success': return styles.bannerSuccess;
      case 'urgent': return styles.bannerUrgent;
      default: return styles.bannerInfo;
    }
  };

  const getIconClass = (type: string): string => {
    switch (type) {
      case 'info': return styles.iconInfo;
      case 'warning': return styles.iconWarning;
      case 'success': return styles.iconSuccess;
      case 'urgent': return styles.iconUrgent;
      default: return styles.iconInfo;
    }
  };

  const getLinkClass = (type: string): string => {
    switch (type) {
      case 'info': return styles.linkInfo;
      case 'warning': return styles.linkWarning;
      case 'success': return styles.linkSuccess;
      case 'urgent': return styles.linkUrgent;
      default: return styles.linkInfo;
    }
  };

  const handleDismiss = (id: string): void => {
    setDismissingId(id);
    window.setTimeout(() => {
      setDismissedIds((currentIds) => [...currentIds, id]);
      setDismissingId(undefined);
    }, 300);
  };

  if (!announcements || announcements.length === 0) {
    return (
      <section className={styles.announcementBannerContainer}>
        <div className={styles.emptyState}>
          <h2>Announcement Banner</h2>
          <p>Add announcements using the property pane.</p>
        </div>
      </section>
    );
  }

  const visibleAnnouncements = announcements.filter(
    (announcement: IAnnouncement) => dismissedIds.indexOf(announcement.uniqueId) === -1
  );

  if (visibleAnnouncements.length === 0) {
    return <section className={styles.announcementBannerContainer} />;
  }

  const containerClass = `${styles.announcementBannerContainer} ${
    stackDirection === 'horizontal' ? styles.announcementBannerContainerHorizontal : ''
  }`;

  return (
    <section className={containerClass}>
      {visibleAnnouncements.map((announcement: IAnnouncement) => {
        const isDismissing = dismissingId === announcement.uniqueId;

        return (
          <div
            key={announcement.uniqueId}
            className={`${styles.banner} ${getBannerClass(announcement.type)} ${isDismissing ? styles.bannerDismissing : ''}`}
          >
            {showIcon && (
              <span className={`${styles.icon} ${getIconClass(announcement.type)}`}>
                {getIcon(announcement.type)}
              </span>
            )}

            <span className={styles.messageText}>
              {escape(announcement.message)}
            </span>

            {announcement.linkUrl && announcement.linkText && (
              <a
                href={announcement.linkUrl}
                className={`${styles.bannerLink} ${getLinkClass(announcement.type)}`}
                target="_blank"
                rel="noreferrer"
              >
                {escape(announcement.linkText)}
              </a>
            )}

            {showDismiss && (
              <button
                className={styles.dismissBtn}
                onClick={() => handleDismiss(announcement.uniqueId)}
                title="Dismiss"
                aria-label="Dismiss announcement"
              >
                ✕
              </button>
            )}
          </div>
        );
      })}
    </section>
  );
};

export default AnnouncementBanner;
