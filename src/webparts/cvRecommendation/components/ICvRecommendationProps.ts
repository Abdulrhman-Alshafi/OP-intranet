import { SPHttpClient } from '@microsoft/sp-http';

export interface ICvRecommendationProps {
  title: string;
  accentColor: string;
  itemsPerPage: number;

  spHttpClient: SPHttpClient;
  siteUrl: string;
  currentUserId: number;
  currentUserName: string;
  isDarkTheme: boolean;
}
