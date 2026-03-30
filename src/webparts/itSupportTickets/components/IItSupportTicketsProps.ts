import { SPHttpClient } from '@microsoft/sp-http';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { ITopTicketService } from '../../../services/ITopTicketService';

export interface IItSupportTicketsProps {
  service: ITopTicketService;
  isConfigured: boolean;
  itAdminGroupName: string;
  currentUserEmail: string;
  currentUserLoginName: string;
  siteUrl: string;
  /** Used for the SharePoint group membership check (isUserInGroup) */
  spHttpClient: SPHttpClient;
  themeVariant?: IReadonlyTheme;
}
