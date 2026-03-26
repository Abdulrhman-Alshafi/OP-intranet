import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ISharePointDocument {
  id: string;
  name: string;
  serverRelativeUrl: string;
  created: string;
  modified: string;
  extension: string;
}

export interface IDepartmentDocuments {
  departmentName: string;
  files: ISharePointDocument[];
}

interface IFilesResponse {
  value: Array<{
    UniqueId: string;
    Name: string;
    ServerRelativeUrl: string;
    TimeCreated: string;
    TimeLastModified: string;
  }>;
}

interface IFoldersResponse {
  value: Array<{
    Name: string;
    ServerRelativeUrl: string;
  }>;
}

/**
 * SharePoint document service.
 * Security is enforced by SharePoint folder/file permissions.
 */
export class SharePointService {
  private readonly _siteUrl: string;
  private readonly _spHttpClient: SPHttpClient;
  private readonly _webServerRelativeUrl: string;

  public constructor(siteUrl: string, webServerRelativeUrl: string, spHttpClient: SPHttpClient) {
    this._siteUrl = siteUrl;
    this._spHttpClient = spHttpClient;
    this._webServerRelativeUrl = webServerRelativeUrl === '/' ? '' : webServerRelativeUrl;
  }

  /**
   * Get personal documents from CompanyDocuments/Personal/{userEmail}
   */
  public async getPersonalFiles(libraryName: string, userEmail: string): Promise<ISharePointDocument[]> {
    const personalFolderPath = this._buildFolderPath(libraryName, `Personal/${userEmail}`);
    const files = await this._getFilesByFolderServerRelativeUrl(personalFolderPath);
    return this._sortFilesByCreatedDate(files);
  }

  /**
   * Get department documents from CompanyDocuments/Departments/*
   * SharePoint returns only folders/files the user can access.
   */
  public async getDepartmentFiles(libraryName: string): Promise<IDepartmentDocuments[]> {
    const departmentsPath = this._buildFolderPath(libraryName, 'Departments');

    const foldersEndpoint = `${this._siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(
      departmentsPath
    )}')/Folders?$select=Name,ServerRelativeUrl`;

    const foldersResponse: SPHttpClientResponse = await this._spHttpClient.get(
      foldersEndpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );

    if (!foldersResponse.ok) {
      throw new Error(`Failed to fetch department folders (${foldersResponse.status}).`);
    }

    const foldersData = (await foldersResponse.json()) as IFoldersResponse;
    const departmentFolders = (foldersData.value || []).filter((folder) => folder.Name !== 'Forms');

    const departmentResults = await Promise.all(
      departmentFolders.map(async (folder) => {
        try {
          const files = await this._getFilesByFolderServerRelativeUrl(folder.ServerRelativeUrl);
          return {
            departmentName: folder.Name,
            files: this._sortFilesByCreatedDate(files)
          } as IDepartmentDocuments;
        } catch {
          // If a specific folder fails, skip it and continue rendering the rest.
          return {
            departmentName: folder.Name,
            files: []
          } as IDepartmentDocuments;
        }
      })
    );

    return departmentResults.filter((department) => department.files.length > 0);
  }

  private async _getFilesByFolderServerRelativeUrl(folderServerRelativeUrl: string): Promise<ISharePointDocument[]> {
    const filesEndpoint = `${this._siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(
      folderServerRelativeUrl
    )}')/Files?$select=UniqueId,Name,ServerRelativeUrl,TimeCreated,TimeLastModified`;

    const response: SPHttpClientResponse = await this._spHttpClient.get(
      filesEndpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );

    if (!response.ok) {
      if (response.status === 404) {
        return [];
      }
      throw new Error(`Failed to fetch files (${response.status}).`);
    }

    const data = (await response.json()) as IFilesResponse;
    return (data.value || []).map((file) => ({
      id: file.UniqueId,
      name: file.Name,
      serverRelativeUrl: file.ServerRelativeUrl,
      created: file.TimeCreated,
      modified: file.TimeLastModified,
      extension: this._getFileExtension(file.Name)
    }));
  }

  private _buildFolderPath(libraryName: string, subPath: string): string {
    return `${this._webServerRelativeUrl}/${libraryName}/${subPath}`.replace(/\/+/g, '/');
  }

  private _sortFilesByCreatedDate(files: ISharePointDocument[]): ISharePointDocument[] {
    return [...files].sort(
      (left, right) => new Date(right.created).getTime() - new Date(left.created).getTime()
    );
  }

  private _getFileExtension(fileName: string): string {
    const lastDotIndex = fileName.lastIndexOf('.');
    if (lastDotIndex === -1 || lastDotIndex === fileName.length - 1) {
      return '';
    }
    return fileName.substring(lastDotIndex + 1).toLowerCase();
  }
}
