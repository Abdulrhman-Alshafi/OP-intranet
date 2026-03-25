import * as React from 'react';
import importedStyles from './AnnouncementBanner.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import type { IAnnouncementBannerProps, IAnnouncement } from './IAnnouncementBannerProps';
import { escape } from '@microsoft/sp-lodash-subset';

interface IAnnouncementBannerState {
  dismissedIds: string[];
  dismissingId: string | undefined;
}

export default class AnnouncementBanner extends React.Component<IAnnouncementBannerProps, IAnnouncementBannerState> {

  constructor(props: IAnnouncementBannerProps) {
    super(props);
    this.state = { dismissedIds: [], dismissingId: undefined };
  }

  private _getIcon(type: string): string {
    switch (type) {
      case 'info': return 'ℹ';
      case 'warning': return '⚠';
      case 'success': return '✓';
      case 'urgent': return '!';
      default: return 'ℹ';
    }
  }

  private _getBannerClass(type: string): string {
    switch (type) {
      case 'info': return styles.bannerInfo;
      case 'warning': return styles.bannerWarning;
      case 'success': return styles.bannerSuccess;
      case 'urgent': return styles.bannerUrgent;
      default: return styles.bannerInfo;
    }
  }

  private _getIconClass(type: string): string {
    switch (type) {
      case 'info': return styles.iconInfo;
      case 'warning': return styles.iconWarning;
      case 'success': return styles.iconSuccess;
      case 'urgent': return styles.iconUrgent;
      default: return styles.iconInfo;
    }
  }

  private _getLinkClass(type: string): string {
    switch (type) {
      case 'info': return styles.linkInfo;
      case 'warning': return styles.linkWarning;
      case 'success': return styles.linkSuccess;
      case 'urgent': return styles.linkUrgent;
      default: return styles.linkInfo;
    }
  }

  private _handleDismiss(id: string): void {
    this.setState({ dismissingId: id });
    // Wait for animation then remove
    setTimeout(() => {
      this.setState((prevState) => ({
        dismissedIds: [...prevState.dismissedIds, id],
        dismissingId: undefined
      }));
    }, 300);
  }

  public render(): React.ReactElement<IAnnouncementBannerProps> {
    const { announcements, showIcon, showDismiss, stackDirection } = this.props;
    const { dismissedIds, dismissingId } = this.state;

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
      (a: IAnnouncement) => dismissedIds.indexOf(a.uniqueId) === -1
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
              className={`${styles.banner} ${this._getBannerClass(announcement.type)} ${isDismissing ? styles.bannerDismissing : ''}`}
            >
              {/* Icon */}
              {showIcon && (
                <span className={`${styles.icon} ${this._getIconClass(announcement.type)}`}>
                  {this._getIcon(announcement.type)}
                </span>
              )}

              {/* Message */}
              <span className={styles.messageText}>
                {escape(announcement.message)}
              </span>

              {/* Link */}
              {announcement.linkUrl && announcement.linkText && (
                <a
                  href={announcement.linkUrl}
                  className={`${styles.bannerLink} ${this._getLinkClass(announcement.type)}`}
                  target="_blank"
                  rel="noreferrer"
                >
                  {escape(announcement.linkText)}
                </a>
              )}

              {/* Dismiss */}
              {showDismiss && (
                <button
                  className={styles.dismissBtn}
                  onClick={() => this._handleDismiss(announcement.uniqueId)}
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
  }
}
