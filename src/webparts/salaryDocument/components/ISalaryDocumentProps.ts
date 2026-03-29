import { MSGraphClientFactory, SPHttpClient } from '@microsoft/sp-http';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SalaryDocumentService } from '../../../services/SalaryDocumentService';

export interface ISalaryDocumentProps {
  service: SalaryDocumentService;
  libraryName: string;
  accountantGroupName: string;
  currentUserEmail: string;
  currentUserLoginName: string;
  /** Needed by @pnp/spfx-controls-react PeoplePicker */
  siteUrl: string;
  spHttpClient: SPHttpClient;
  msGraphClientFactory: MSGraphClientFactory;
  themeVariant?: IReadonlyTheme;
}
