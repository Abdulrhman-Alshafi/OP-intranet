import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISalesPerformanceDashboardProps {
  listId: string;
  chartType: string;
  isDarkTheme: boolean;
  context: WebPartContext;
}
