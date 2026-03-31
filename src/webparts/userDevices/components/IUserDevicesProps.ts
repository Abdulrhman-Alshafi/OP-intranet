import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { HelpDeskDeviceService } from '../../../services/HelpDeskDeviceService';

export interface IUserDevicesProps {
  service: HelpDeskDeviceService;
  devicesListName: string;
  requestsListName: string;
  themeVariant: IReadonlyTheme | undefined;
}
