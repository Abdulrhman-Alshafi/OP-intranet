import { MSGraphClientFactory, SPHttpClient } from '@microsoft/sp-http';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { HrDocumentService } from '../../../services/HrDocumentService';

export interface IHrDocumentProps {
  service: HrDocumentService;
  libraryName: string;
  hrGroupName: string;
  currentUserEmail: string;
  currentUserLoginName: string;
  /** Needed by @pnp/spfx-controls-react PeoplePicker */
  siteUrl: string;
  spHttpClient: SPHttpClient;
  msGraphClientFactory: MSGraphClientFactory;
  themeVariant?: IReadonlyTheme;
}
