import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { HelpDeskDeviceService } from '../../../services/HelpDeskDeviceService';

export interface IHelpDeskDevicesProps {
  service: HelpDeskDeviceService;
  devicesListName: string;
  requestsListName: string;
  helpDeskGroupName: string;
  themeVariant: IReadonlyTheme | undefined;
}
