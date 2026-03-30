import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { FAQService } from '../../../services/FAQService';

export interface IFAQSectionProps {
  service: FAQService;
  themeVariant?: IReadonlyTheme;
}
