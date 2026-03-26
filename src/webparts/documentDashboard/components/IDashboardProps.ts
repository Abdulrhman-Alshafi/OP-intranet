import { SPHttpClient } from '@microsoft/sp-http';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { IDepartmentDocuments, ISharePointDocument } from '../../../services/sharepointService';

export interface IDashboardProps {
  title: string;
  libraryName: string;
  userEmail: string;
  siteUrl: string;
  webServerRelativeUrl: string;
  spHttpClient: SPHttpClient;
  themeVariant?: IReadonlyTheme;
}

export interface IPersonalDocumentsProps {
  files: ISharePointDocument[];
  loading: boolean;
  error?: string;
  siteUrl: string;
}

export interface IDepartmentDocumentsProps {
  departmentFiles: IDepartmentDocuments[];
  loading: boolean;
  error?: string;
  siteUrl: string;
}
