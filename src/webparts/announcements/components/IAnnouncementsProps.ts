import { SPHttpClient } from '@microsoft/sp-http';

export interface IAnnouncementsProps {
  // WebPart title
  title: string;

  // Property pane config
  announcementsListName: string;
  enableCreateAnnouncement: boolean;
  enableReactions: boolean;
  enableComments: boolean;
  itemsPerPage: number;

  // Context
  spHttpClient: SPHttpClient;
  siteUrl: string;
  userEmail: string;
  userId: number;

  // Theme
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
}
