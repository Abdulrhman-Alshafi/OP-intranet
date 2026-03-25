import { SPHttpClient } from '@microsoft/sp-http';

export interface IListViewerProps {
  listId: string;
  maxItems: number;
  viewStyle: 'grid' | 'list' | 'table';
  title: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  hasTeamsContext: boolean;
  themeVariant: any;
}
