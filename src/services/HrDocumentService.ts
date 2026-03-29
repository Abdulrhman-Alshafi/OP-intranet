import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ISharePointDocument } from './sharepointService';

// ─── Public Interfaces ────────────────────────────────────────────────────────

export interface IHrDocument extends ISharePointDocument {
  employeeEmail: string;
  employeeDisplayName: string;
  /** SharePoint list item integer ID – used for permission assignment */
  listItemId: number;
  /** SP internal integer user ID of the Employee Person/Group field */
  employeeId: number;
  /** Document type, e.g. "Offer Letter", "Contract", "Warning Letter" */
  documentType: string;
}

export interface IHrUploadJob {
  file: File;
  /** User's loginName, e.g. i:0#.f|membership|user@tenant.onmicrosoft.com */
  employeeLoginName: string;
  /** SP integer user ID – resolve once before queueing so permissions are fast */
  employeeId: number;
  /** Document type string */
  documentType: string;
}

export type HrUploadStatus = 'pending' | 'uploading' | 'success' | 'error';

export interface IHrUploadResult {
  fileName: string;
  status: HrUploadStatus;
  message?: string;
}

export interface IPagedHrDocuments {
  items: IHrDocument[];
  /** Pass this back to get the next page; undefined means last page. */
  nextLink: string | undefined;
}

// ─── Internal SP Response Shapes ─────────────────────────────────────────────

interface ISPItemResponse {
  value: ISPItem[];
  'odata.nextLink'?: string;
}

interface ISPItem {
  ID: number;
  Employee?: {
    Id: number;
    EMail: string;
    Title: string;
  };
  DocumentType?: string;
  File: {
    UniqueId: string;
    Name: string;
    ServerRelativeUrl: string;
    TimeCreated: string;
    TimeLastModified: string;
  };
}

interface ISPFileAddResponse {
  ListItemAllFields: {
    ID: number;
  };
}

// ─── Concurrency helper ───────────────────────────────────────────────────────

async function runInPool<T>(
  tasks: (() => Promise<T>)[],
  poolSize: number
): Promise<T[]> {
  const results: T[] = new Array(tasks.length);
  let nextIndex = 0;

  async function worker(): Promise<void> {
    while (nextIndex < tasks.length) {
      const i = nextIndex++;
      results[i] = await tasks[i]();
    }
  }

  await Promise.all(Array.from({ length: Math.min(poolSize, tasks.length) }, worker));
  return results;
}

// ─── Service Class ────────────────────────────────────────────────────────────

/**
 * Encapsulates all SharePoint REST calls for the HR Document web part.
 * Security is enforced at item level – each uploaded document has inheritance
 * broken and only the target employee (Read) and the HR group (Edit)
 * are granted access.
 */
export class HrDocumentService {
  private readonly _spHttpClient: SPHttpClient;
  private readonly _siteUrl: string;

  /** Cached SP integer ID for the HR group – resolved on first call. */
  private _hrGroupId: number | undefined;

  public constructor(siteUrl: string, spHttpClient: SPHttpClient) {
    this._siteUrl = siteUrl;
    this._spHttpClient = spHttpClient;
  }

  // ── List-item queries ─────────────────────────────────────────────────────

  /**
   * Returns HR documents where the Employee field matches the given email.
   * Used for the employee-facing view (own documents only).
   */
  public async getHrDocuments(
    libraryName: string,
    userEmail: string
  ): Promise<IHrDocument[]> {
    const encoded = encodeURIComponent(libraryName);
    const filter = `Employee/EMail eq '${encodeURIComponent(userEmail)}'`;

    const url =
      `${this._siteUrl}/_api/web/lists/getbytitle('${encoded}')/items` +
      `?$filter=${filter}` +
      `&$expand=Employee,File` +
      `&$select=ID,Employee/Id,Employee/EMail,Employee/Title,DocumentType,File/UniqueId,File/Name,File/ServerRelativeUrl,File/TimeCreated,File/TimeLastModified` +
      `&$orderby=Created desc` +
      `&$top=100`;

    const response = await this._get(url);
    const data = (await response.json()) as ISPItemResponse;
    return (data.value || []).map(this._mapItem);
  }

  /**
   * Returns a page of all HR documents across all employees.
   * Used for the HR-facing view.
   */
  public async getAllHrDocuments(
    libraryName: string,
    pageSize: number,
    nextLink?: string
  ): Promise<IPagedHrDocuments> {
    const encoded = encodeURIComponent(libraryName);

    const url =
      nextLink ||
      `${this._siteUrl}/_api/web/lists/getbytitle('${encoded}')/items` +
        `?$expand=Employee,File` +
        `&$select=ID,Employee/Id,Employee/EMail,Employee/Title,DocumentType,File/UniqueId,File/Name,File/ServerRelativeUrl,File/TimeCreated,File/TimeLastModified` +
        `&$orderby=Created desc` +
        `&$top=${pageSize}`;

    const response = await this._get(url);
    const data = (await response.json()) as ISPItemResponse;

    return {
      items: (data.value || []).map(this._mapItem),
      nextLink: data['odata.nextLink'] ?? undefined
    };
  }

  // ── Group membership ──────────────────────────────────────────────────────

  /**
   * Checks whether the current user is a member of the named SP group.
   */
  public async isUserInGroup(groupName: string): Promise<boolean> {
    const url = `${this._siteUrl}/_api/web/currentuser/groups?$select=Title`;
    const response = await this._get(url);
    const data = (await response.json()) as { value: { Title: string }[] };
    return (data.value || []).some((g) => g.Title === groupName);
  }

  /**
   * Resolves the integer SP group ID for the HR group and caches it.
   * Throws if the group does not exist.
   */
  public async resolveHrGroupId(groupName: string): Promise<number> {
    if (this._hrGroupId !== undefined) {
      return this._hrGroupId;
    }
    const encoded = encodeURIComponent(groupName);
    const url = `${this._siteUrl}/_api/web/sitegroups/getbyname('${encoded}')?$select=Id`;
    const response = await this._get(url);
    if (!response.ok) {
      throw new Error(
        `HR group "${groupName}" not found (HTTP ${response.status}).`
      );
    }
    const data = (await response.json()) as { Id: number };
    this._hrGroupId = data.Id;
    return this._hrGroupId;
  }

  /**
   * Resolves a user's SP integer ID and loginName from their email address.
   */
  public async resolveUser(
    email: string
  ): Promise<{ id: number; loginName: string; displayName: string } | undefined> {
    const url = `${this._siteUrl}/_api/web/ensureuser('${encodeURIComponent(email)}')`;
    const response = await this._post(url, {});
    if (!response.ok) return undefined;
    const data = (await response.json()) as { Id: number; LoginName: string; Title: string };
    return { id: data.Id, loginName: data.LoginName, displayName: data.Title };
  }

  // ── Upload ────────────────────────────────────────────────────────────────

  /**
   * Full 4-step upload flow:
   *  1. Upload file binary to library root
   *  2. Set Employee + DocumentType fields on the list item
   *  3. Break role inheritance (no copy, clear subscopes)
   *  4. Assign Employee → Read, HR group → Edit in parallel
   */
  public async uploadHrDocument(
    libraryName: string,
    file: File,
    employeeLoginName: string,
    employeeId: number,
    hrGroupId: number,
    documentType: string = ''
  ): Promise<void> {
    const encodedLib = encodeURIComponent(libraryName);
    const encodedName = encodeURIComponent(file.name);

    // ── Step 1: Upload file ──────────────────────────────────────────────
    const uploadUrl =
      `${this._siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodedLib}')` +
      `/Files/add(url='${encodedName}',overwrite=true)` +
      `?$expand=ListItemAllFields`;

    const arrayBuffer = await file.arrayBuffer();

    const uploadResponse = await this._spHttpClient.post(
      uploadUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/octet-stream',
          'odata-version': ''
        },
        body: arrayBuffer
      }
    );

    if (!uploadResponse.ok) {
      throw new Error(`Upload failed (HTTP ${uploadResponse.status}): ${file.name}`);
    }

    const uploadData = (await uploadResponse.json()) as ISPFileAddResponse;
    const itemId = uploadData.ListItemAllFields?.ID;
    if (!itemId) {
      throw new Error(`Could not get item ID after uploading "${file.name}".`);
    }

    // ── Step 2: Set Employee + DocumentType fields ────────────────────────
    const itemUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${encodedLib}')/items(${itemId})`;
    const patchBody: { EmployeeId: number; DocumentType?: string } = { EmployeeId: employeeId };
    if (documentType) patchBody.DocumentType = documentType;

    const patchResponse = await this._spHttpClient.post(itemUrl, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify(patchBody)
    });

    if (!patchResponse.ok) {
      throw new Error(
        `Failed to set Employee/DocumentType fields on item ${itemId} (HTTP ${patchResponse.status}).`
      );
    }

    // ── Step 3: Break role inheritance ───────────────────────────────────
    const breakUrl =
      `${this._siteUrl}/_api/web/lists/getbytitle('${encodedLib}')/items(${itemId})` +
      `/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`;

    const breakResponse = await this._post(breakUrl, {});
    if (!breakResponse.ok) {
      throw new Error(`Failed to break role inheritance on item ${itemId}.`);
    }

    // ── Step 4: Grant permissions (parallel) ─────────────────────────────
    const basePermUrl =
      `${this._siteUrl}/_api/web/lists/getbytitle('${encodedLib}')/items(${itemId})` +
      `/roleassignments`;

    const [empResponse, grpResponse] = await Promise.all([
      this._post(
        `${basePermUrl}/addroleassignment(principalid=${employeeId},roledefid=1073741826)`,
        {}
      ),
      this._post(
        `${basePermUrl}/addroleassignment(principalid=${hrGroupId},roledefid=1073741827)`,
        {}
      )
    ]);

    if (!empResponse.ok) {
      throw new Error(`Failed to grant Read to employee on item ${itemId}.`);
    }
    if (!grpResponse.ok) {
      throw new Error(`Failed to grant Edit to HR group on item ${itemId}.`);
    }
  }

  /**
   * Upload multiple documents with a concurrency limit of 3.
   * Calls onProgress after each file completes/fails.
   */
  public async uploadBulkHrDocuments(
    libraryName: string,
    jobs: IHrUploadJob[],
    hrGroupId: number,
    onProgress: (result: IHrUploadResult) => void
  ): Promise<void> {
    const tasks = jobs.map((job) => async (): Promise<void> => {
      try {
        await this.uploadHrDocument(
          libraryName,
          job.file,
          job.employeeLoginName,
          job.employeeId,
          hrGroupId,
          job.documentType
        );
        onProgress({ fileName: job.file.name, status: 'success' });
      } catch (err) {
        onProgress({
          fileName: job.file.name,
          status: 'error',
          message: err instanceof Error ? err.message : 'Unknown error'
        });
      }
    });

    await runInPool(tasks, 3);
  }

  /**
   * Returns true if a file with the given name already exists in the library.
   */
  public async fileExists(libraryName: string, fileName: string): Promise<boolean> {
    const encodedLib = encodeURIComponent(libraryName);
    const encodedName = encodeURIComponent(fileName.replace(/'/g, "''"));
    const url =
      `${this._siteUrl}/_api/web/lists/getbytitle('${encodedLib}')/items` +
      `?$filter=FileLeafRef eq '${encodedName}'` +
      `&$select=ID&$top=1`;
    const response = await this._get(url);
    if (!response.ok) return false;
    const data = (await response.json()) as { value: unknown[] };
    return data.value.length > 0;
  }

  /**
   * Permanently deletes a library item (the file + its list item).
   */
  public async deleteDocument(libraryName: string, itemId: number): Promise<void> {
    const encodedLib = encodeURIComponent(libraryName);
    const url =
      `${this._siteUrl}/_api/web/lists/getbytitle('${encodedLib}')/items(${itemId})`;

    const response = await this._spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'odata-version': '',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'DELETE'
        },
        body: JSON.stringify({})
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to delete document (HTTP ${response.status}).`);
    }
  }

  /**
   * Returns true when the current user is a site collection administrator.
   */
  public async isCurrentUserSiteAdmin(): Promise<boolean> {
    const url = `${this._siteUrl}/_api/web/currentuser?$select=IsSiteAdmin`;
    const response = await this._get(url);
    if (!response.ok) return false;
    const data = (await response.json()) as { IsSiteAdmin?: boolean };
    return data.IsSiteAdmin === true;
  }

  // ── Private helpers ───────────────────────────────────────────────────────

  private async _get(url: string): Promise<SPHttpClientResponse> {
    return this._spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'odata-version': ''
      }
    });
  }

  private async _post(url: string, body: object): Promise<SPHttpClientResponse> {
    return this._spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify(body)
    });
  }

  private readonly _mapItem = (item: ISPItem): IHrDocument => {
    const name = item.File?.Name ?? '';
    const dotIdx = name.lastIndexOf('.');
    return {
      id: item.File?.UniqueId ?? String(item.ID),
      listItemId: item.ID,
      name,
      serverRelativeUrl: item.File?.ServerRelativeUrl ?? '',
      created: item.File?.TimeCreated ?? '',
      modified: item.File?.TimeLastModified ?? '',
      extension: dotIdx !== -1 ? name.slice(dotIdx + 1).toLowerCase() : '',
      employeeEmail: item.Employee?.EMail ?? '',
      employeeDisplayName: item.Employee?.Title ?? '',
      employeeId: item.Employee?.Id ?? 0,
      documentType: item.DocumentType ?? ''
    };
  };
}
