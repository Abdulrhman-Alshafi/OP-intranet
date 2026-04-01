import { SPHttpClient } from '@microsoft/sp-http';

export interface IAnnouncement {
  Id: number;
  Title: string;
  Description: string;
  Category: string;
  ImageUrl: string;
  IsImportant: boolean;
  Created: string;
  IsDismissed?: boolean;
}

export interface IAnnouncementsHubProps {
  title: string;
  layoutMode: 'grid' | 'list';
  itemsToDisplay: number;
  sortOrder: 'newest' | 'oldest';
  showImages: boolean;
  enableAnimations: boolean;
  enableCategoryColors: boolean;
  accentColor: string;

  spHttpClient: SPHttpClient;
  siteUrl: string;
  currentUserId: number;
  isAdmin: boolean;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
}
