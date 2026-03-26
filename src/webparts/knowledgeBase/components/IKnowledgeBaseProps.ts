import { SPHttpClient } from '@microsoft/sp-http';

export interface IKnowledgeBaseProps {
  // WebPart title
  title: string;

  // Property pane config
  listName: string;
  defaultCategory: string;
  itemsPerPage: number;
  enableSearch: boolean;
  enableCategoryFilter: boolean;

  // Context
  spHttpClient: SPHttpClient;
  siteUrl: string;

  // Theme
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
