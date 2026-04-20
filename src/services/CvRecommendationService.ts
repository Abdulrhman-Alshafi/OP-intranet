import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';

// ─── Interfaces ─────────────────────────────────────────────────────────────

export interface ICvRecommendation {
  Id: number;
  Title: string;                  // Candidate Name
  CandidateEmail: string;
  PhoneNumber: string;
  Position: string;
  Notes: string;
  Status: 'Submitted' | 'Under Review' | 'Accepted' | 'Rejected';
  SubmittedById: number;
  SubmittedByName: string;
  Created: string;
  CVFileUrl?: string;             // URL after upload to document library
  CVFileName?: string;
}

export interface ICvStats {
  total: number;
  submitted: number;
  underReview: number;
  accepted: number;
  rejected: number;
}

// ─── Constants ───────────────────────────────────────────────────────────────

const LIST_NAME = 'CVRecommendations';
const LIBRARY_NAME = 'CVRecommendationFiles';
const HR_GROUP_NAME = 'HR Team';

/**
 * Service class for all SharePoint operations related to CV Recommendations.
 * Uses SPHttpClient following the same pattern as other services in this project.
 */
export class CvRecommendationService {
  private _spHttpClient: SPHttpClient;
  private _siteUrl: string;
  private _listsEnsured = false;

  constructor(spHttpClient: SPHttpClient, siteUrl: string) {
    this._spHttpClient = spHttpClient;
    this._siteUrl = siteUrl;
  }

  // ─── List / Library Provisioning ─────────────────────────────────────────

  /**
   * Ensure the CV Recommendations list and CV files library exist.
   * Creates them if missing.
   */
  public async ensureLists(): Promise<void> {
    if (this._listsEnsured) return;
    try {
      await this._ensureCvList();
      await this._ensureCvLibrary();
      this._listsEnsured = true;
    } catch (error) {
      console.error('CvRecommendationService: ensureLists error', error);
      this._listsEnsured = true;
    }
  }

  private async _ensureCvList(): Promise<void> {
    const checkUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_NAME}')`;
    const checkRes = await this._spHttpClient.get(checkUrl, SPHttpClient.configurations.v1);
    if (checkRes.ok) return;

    // Create list
    const createRes = await this._spHttpClient.post(
      `${this._siteUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1,
      {
        body: JSON.stringify({
          Title: LIST_NAME,
          Description: 'Stores CV Recommendations submitted by employees',
          BaseTemplate: 100,
          AllowContentTypes: false,
          ContentTypesEnabled: false,
          EnableAttachments: false
        })
      } as ISPHttpClientOptions
    );
    if (!createRes.ok) {
      console.warn(`CvRecommendationService: Could not create list "${LIST_NAME}"`);
      return;
    }

    // Add custom columns
    const fields: Array<{ type: string; title: string; choices?: string[] }> = [
      { type: 'Text', title: 'CandidateEmail' },
      { type: 'Text', title: 'PhoneNumber' },
      { type: 'Text', title: 'Position' },
      { type: 'Note', title: 'Notes' },
      {
        type: 'Choice', title: 'Status',
        choices: ['Submitted', 'Under Review', 'Accepted', 'Rejected']
      },
      { type: 'Text', title: 'CVFileUrl' },
      { type: 'Text', title: 'CVFileName' }
    ];

    for (const f of fields) {
      const fieldBody: Record<string, unknown> = {
        Title: f.title,
        Required: false
      };
      if (f.type === 'Choice') {
        fieldBody['__metadata'] = { type: 'SP.FieldChoice' };
        fieldBody['FieldTypeKind'] = 6;
        fieldBody['Choices'] = { results: f.choices };
        fieldBody['DefaultValue'] = f.choices?.[0] ?? '';
      } else if (f.type === 'Note') {
        fieldBody['__metadata'] = { type: 'SP.FieldMultiLineText' };
        fieldBody['FieldTypeKind'] = 3;
      } else {
        fieldBody['__metadata'] = { type: 'SP.Field' };
        fieldBody['FieldTypeKind'] = 2;
      }

      await this._spHttpClient.post(
        `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_NAME}')/fields`,
        SPHttpClient.configurations.v1,
        { body: JSON.stringify(fieldBody) } as ISPHttpClientOptions
      );
    }
  }

  private async _ensureCvLibrary(): Promise<void> {
    const checkUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')`;
    const checkRes = await this._spHttpClient.get(checkUrl, SPHttpClient.configurations.v1);
    if (checkRes.ok) return;

    await this._spHttpClient.post(
      `${this._siteUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1,
      {
        body: JSON.stringify({
          Title: LIBRARY_NAME,
          Description: 'Stores CV files uploaded as part of recommendations',
          BaseTemplate: 101   // Document Library
        })
      } as ISPHttpClientOptions
    );
  }

  // ─── Current User ─────────────────────────────────────────────────────────

  /**
   * Resolve whether the current user belongs to the HR Team group.
   */
  public async isHrUser(): Promise<boolean> {
    try {
      const url = `${this._siteUrl}/_api/web/currentuser/groups?$select=Title`;
      const res = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!res.ok) return false;
      const json = await res.json();
      const groups: { Title: string }[] = json?.value ?? [];
      return groups.some(g => g.Title === HR_GROUP_NAME);
    } catch {
      return false;
    }
  }

  // ─── CRUD ─────────────────────────────────────────────────────────────────

  /**
   * Fetch CV recommendations. HR sees all; employees see only their own.
   */
  public async getCvRecommendations(
    currentUserId: number,
    isHr: boolean,
    filterStatus?: string,
    filterPosition?: string,
    searchText?: string
  ): Promise<ICvRecommendation[]> {
    const select = '$select=Id,Title,CandidateEmail,PhoneNumber,Position,Notes,Status,CVFileUrl,CVFileName,Created,Author/Id,Author/Title';
    const expand = '$expand=Author';
    const orderby = '$orderby=Created desc';

    let filters: string[] = [];
    if (!isHr) {
      filters.push(`AuthorId eq ${currentUserId}`);
    }
    if (filterStatus && filterStatus !== 'All') {
      filters.push(`Status eq '${filterStatus}'`);
    }
    if (filterPosition && filterPosition !== 'All') {
      filters.push(`substringof('${filterPosition}',Position)`);
    }
    if (searchText) {
      const safe = searchText.replace(/'/g, "''");
      filters.push(`(substringof('${safe}',Title) or substringof('${safe}',CandidateEmail) or substringof('${safe}',Position))`);
    }

    const filter = filters.length ? `$filter=${filters.join(' and ')}` : '';
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_NAME}')/items?${select}&${expand}&${orderby}${filter ? '&' + filter : ''}`;

    const res = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err?.error?.message ?? 'Failed to load CV recommendations.');
    }
    const json = await res.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (json.value ?? []).map((item: any): ICvRecommendation => ({
      Id: item.Id,
      Title: item.Title ?? '',
      CandidateEmail: item.CandidateEmail ?? '',
      PhoneNumber: item.PhoneNumber ?? '',
      Position: item.Position ?? '',
      Notes: item.Notes ?? '',
      Status: item.Status ?? 'Submitted',
      SubmittedById: item.Author?.Id ?? 0,
      SubmittedByName: item.Author?.Title ?? '',
      Created: item.Created ?? '',
      CVFileUrl: item.CVFileUrl ?? '',
      CVFileName: item.CVFileName ?? ''
    }));
  }

  /**
   * Submit a new CV recommendation (with optional file upload).
   */
  public async submitCvRecommendation(
    candidateName: string,
    candidateEmail: string,
    phoneNumber: string,
    position: string,
    notes: string,
    file?: File
  ): Promise<void> {
    let fileUrl = '';
    let fileName = '';

    // Upload file first if provided
    if (file) {
      const result = await this._uploadFile(file);
      fileUrl = result.url;
      fileName = result.name;
    }

    const body = JSON.stringify({
      Title: candidateName,
      CandidateEmail: candidateEmail,
      PhoneNumber: phoneNumber,
      Position: position,
      Notes: notes,
      Status: 'Submitted',
      CVFileUrl: fileUrl,
      CVFileName: fileName
    });

    const res = await this._spHttpClient.post(
      `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_NAME}')/items`,
      SPHttpClient.configurations.v1,
      {
        body,
        headers: {
          'Content-Type': 'application/json;odata=verbose',
          Accept: 'application/json;odata=verbose'
        }
      } as ISPHttpClientOptions
    );
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err?.error?.message ?? 'Failed to submit CV recommendation.');
    }
  }

  /**
   * Update the status of a CV (HR only).
   */
  public async updateStatus(
    itemId: number,
    newStatus: ICvRecommendation['Status']
  ): Promise<void> {
    const res = await this._spHttpClient.post(
      `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_NAME}')/items(${itemId})`,
      SPHttpClient.configurations.v1,
      {
        body: JSON.stringify({ Status: newStatus }),
        headers: {
          'Content-Type': 'application/json;odata=verbose',
          Accept: 'application/json;odata=verbose',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        }
      } as ISPHttpClientOptions
    );
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err?.error?.message ?? 'Failed to update status.');
    }
  }

  /**
   * Delete a CV recommendation.
   */
  public async deleteCvRecommendation(itemId: number): Promise<void> {
    const res = await this._spHttpClient.post(
      `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_NAME}')/items(${itemId})`,
      SPHttpClient.configurations.v1,
      {
        body: JSON.stringify({}),
        headers: {
          'Content-Type': 'application/json;odata=verbose',
          Accept: 'application/json;odata=verbose',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'DELETE'
        }
      } as ISPHttpClientOptions
    );
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err?.error?.message ?? 'Failed to delete CV recommendation.');
    }
  }

  /**
   * Compute stats across items.
   */
  public computeStats(items: ICvRecommendation[]): ICvStats {
    return {
      total: items.length,
      submitted: items.filter(i => i.Status === 'Submitted').length,
      underReview: items.filter(i => i.Status === 'Under Review').length,
      accepted: items.filter(i => i.Status === 'Accepted').length,
      rejected: items.filter(i => i.Status === 'Rejected').length
    };
  }

  // ─── File Upload ──────────────────────────────────────────────────────────

  private async _uploadFile(file: File): Promise<{ url: string; name: string }> {
    const safeName = file.name.replace(/['"]/g, '_');
    const uploadUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${LIBRARY_NAME}')/rootfolder/files/add(url='${safeName}',overwrite=true)`;

    const arrayBuffer = await file.arrayBuffer();
    const res = await this._spHttpClient.post(uploadUrl, SPHttpClient.configurations.v1, {
      body: arrayBuffer,
      headers: {
        Accept: 'application/json;odata=verbose'
      }
    } as ISPHttpClientOptions);

    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err?.error?.message ?? 'Failed to upload file.');
    }

    const json = await res.json();
    const serverRelativeUrl: string = json?.d?.ServerRelativeUrl ?? '';
    return {
      url: serverRelativeUrl ? `${this._siteUrl.split('/_api')[0]}${serverRelativeUrl}` : '',
      name: safeName
    };
  }
}
