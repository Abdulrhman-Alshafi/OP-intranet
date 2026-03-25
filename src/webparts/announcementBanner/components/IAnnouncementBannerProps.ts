export interface IAnnouncement {
  uniqueId: string;
  message: string;
  type: 'info' | 'warning' | 'success' | 'urgent';
  linkText?: string;
  linkUrl?: string;
}

export interface IAnnouncementBannerProps {
  announcements: IAnnouncement[];
  showIcon: boolean;
  showDismiss: boolean;
  stackDirection: 'vertical' | 'horizontal';
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
}
