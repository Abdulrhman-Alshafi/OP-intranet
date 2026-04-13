import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';

/**
 * Interface for an Announcement item from SharePoint.
 */
export interface IAnnouncement {
  Id: number;
  Title: string;
  Description: string;
  Category: string;
  ImageUrl: string;
  IsImportant: boolean;
  Created: string;
  IsDismissed?: boolean;
}

/**
 * List names used by the Announcements Hub.
 */
const LIST_ANNOUNCEMENTS = 'OPAnnouncements';
const LIST_DISMISSED = 'OPAnnouncementsDismissed';

/**
 * Service class to handle all SharePoint list operations for the Announcements Hub.
 * Uses SPHttpClient following the same pattern as RecognitionService.
 */
export class AnnouncementsService {
  private _spHttpClient: SPHttpClient;
  private _siteUrl: string;
  private _listsEnsured: boolean = false;

  constructor(spHttpClient: SPHttpClient, siteUrl: string) {
    this._spHttpClient = spHttpClient;
    this._siteUrl = siteUrl;
  }

  // ─── List Provisioning ───────────────────────────────────────────────

  /**
   * Ensure all required SharePoint lists exist. Creates them if missing.
   */
  public async ensureLists(): Promise<void> {
    if (this._listsEnsured) return;

    try {
      await this._ensureList(LIST_ANNOUNCEMENTS, 'Stores organizational announcements', [
        { type: 'Note', title: 'Description' },
        {
          type: 'Choice', title: 'Category', choices: [
            'General', 'HR', 'IT', 'Finance', 'Events', 'Policy', 'Urgent', 'Press release'
          ]
        },
        { type: 'Text', title: 'ImageUrl' },
        { type: 'Boolean', title: 'IsImportant' }
      ]);

      await this._ensureList(LIST_DISMISSED, 'Tracks dismissed announcements per user', [
        { type: 'Number', title: 'AnnouncementId' },
        { type: 'Number', title: 'UserId' }
      ]);

      this._listsEnsured = true;
    } catch (error) {
      console.error('AnnouncementsService: Failed to ensure lists', error);
      this._listsEnsured = true;
    }
  }

  /**
   * Create a list if it doesn't already exist.
   */
  private async _ensureList(
    listTitle: string,
    description: string,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    additionalFields: any[]
  ): Promise<void> {
    const checkUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${listTitle}')`;
    const checkResponse = await this._spHttpClient.get(checkUrl, SPHttpClient.configurations.v1);
    if (checkResponse.ok) return;

    const createUrl = `${this._siteUrl}/_api/web/lists`;
    const createOptions: ISPHttpClientOptions = {
      body: JSON.stringify({
        Title: listTitle,
        Description: description,
        BaseTemplate: 100,
        AllowContentTypes: false,
        ContentTypesEnabled: false
      })
    };

    const createResponse = await this._spHttpClient.post(createUrl, SPHttpClient.configurations.v1, createOptions);
    if (!createResponse.ok) {
      const err = await createResponse.json().catch(() => ({}));
      console.warn(`AnnouncementsService: Could not create list "${listTitle}"`, err);
      return;
    }

    for (const field of additionalFields) {
      let schemaXml = '';
      if (field.type === 'Note') {
        schemaXml = `<Field Type="Note" DisplayName="${field.title}" Name="${field.title}" StaticName="${field.title}" />`;
      } else if (field.type === 'Text') {
        schemaXml = `<Field Type="Text" DisplayName="${field.title}" Name="${field.title}" StaticName="${field.title}" />`;
      } else if (field.type === 'Number') {
        schemaXml = `<Field Type="Number" DisplayName="${field.title}" Name="${field.title}" StaticName="${field.title}" />`;
      } else if (field.type === 'Boolean') {
        schemaXml = `<Field Type="Boolean" DisplayName="${field.title}" Name="${field.title}" StaticName="${field.title}"><Default>0</Default></Field>`;
      } else if (field.type === 'Choice') {
        const choicesXml = field.choices.map((c: string) => `<CHOICE>${c.replace(/&/g, '&amp;')}</CHOICE>`).join('');
        schemaXml = `<Field Type="Choice" DisplayName="${field.title}" Name="${field.title}" StaticName="${field.title}"><CHOICES>${choicesXml}</CHOICES></Field>`;
      }

      if (schemaXml) {
        const addFieldUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${listTitle}')/fields/createfieldasxml`;
        const addFieldBody = JSON.stringify({
          parameters: {
            "__metadata": { "type": "SP.XmlSchemaFieldCreationInformation" },
            SchemaXml: schemaXml,
            Options: 8
          }
        });

        await this._spHttpClient.post(addFieldUrl, SPHttpClient.configurations.v1, {
          headers: {
            'Content-Type': 'application/json;odata=verbose',
            'Accept': 'application/json;odata=verbose'
          },
          body: addFieldBody
        });
      }
    }
  }

  // ─── Announcements CRUD ──────────────────────────────────────────────

  /**
   * Fetch announcements with optional sorting.
   */
  public async getAnnouncements(
    top: number = 10,
    sortOrder: 'newest' | 'oldest' = 'newest',
    currentUserId?: number
  ): Promise<IAnnouncement[]> {
    await this.ensureLists();

    const orderBy = sortOrder === 'newest' ? 'Created desc' : 'Created asc';
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_ANNOUNCEMENTS}')/items` +
      `?$select=Id,Title,Description,Category,ImageUrl,IsImportant,Created` +
      `&$orderby=${orderBy}` +
      `&$top=${top}`;

    const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!response.ok) {
      throw new Error(`Failed to fetch announcements: ${response.statusText}`);
    }

    const json = await response.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const announcements: IAnnouncement[] = (json.value || []).map((item: any) => ({
      Id: item.Id,
      Title: item.Title || '',
      Description: item.Description || '',
      Category: item.Category || 'General',
      ImageUrl: item.ImageUrl || '',
      IsImportant: !!item.IsImportant,
      Created: item.Created,
      IsDismissed: false
    }));

    // Enrich with dismissal status
    if (currentUserId && currentUserId > 0) {
      await this._enrichWithDismissals(announcements, currentUserId);
    }

    return announcements;
  }

  /**
   * Enrich announcements with user-specific dismissal status.
   */
  private async _enrichWithDismissals(announcements: IAnnouncement[], userId: number): Promise<void> {
    if (announcements.length === 0) return;

    try {
      const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_DISMISSED}')/items` +
        `?$select=AnnouncementId&$filter=UserId eq ${userId}&$top=5000`;
      const response = await this._spHttpClient.get(url, SPHttpClient.configurations.v1);
      if (!response.ok) return;

      const json = await response.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const dismissedIds = new Set((json.value || []).map((d: any) => d.AnnouncementId));

      announcements.forEach(a => {
        a.IsDismissed = dismissedIds.has(a.Id);
      });
    } catch (error) {
      console.warn('AnnouncementsService: Could not fetch dismissals', error);
    }
  }

  /**
   * Create a new announcement (admin only).
   */
  public async createAnnouncement(
    title: string,
    description: string,
    category: string,
    imageUrl: string,
    isImportant: boolean
  ): Promise<void> {
    await this.ensureLists();

    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_ANNOUNCEMENTS}')/items`;
    const body = JSON.stringify({
      Title: title,
      Description: description,
      Category: category,
      ImageUrl: imageUrl || '',
      IsImportant: isImportant
    });

    const response = await this._spHttpClient.post(url, SPHttpClient.configurations.v1, { body });
    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(`Failed to create announcement: ${JSON.stringify(err)}`);
    }
  }

  /**
   * Delete an announcement (admin only).
   */
  public async deleteAnnouncement(announcementId: number): Promise<void> {
    // Delete related dismissals first
    const dismissUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_DISMISSED}')/items` +
      `?$select=Id&$filter=AnnouncementId eq ${announcementId}`;
    const dismissResp = await this._spHttpClient.get(dismissUrl, SPHttpClient.configurations.v1);
    if (dismissResp.ok) {
      const dismissJson = await dismissResp.json();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      for (const d of (dismissJson.value || [])) {
        await this._deleteItem(LIST_DISMISSED, d.Id);
      }
    }

    await this._deleteItem(LIST_ANNOUNCEMENTS, announcementId);
  }

  /**
   * Dismiss an announcement for the current user.
   */
  public async dismissAnnouncement(announcementId: number, userId: number): Promise<void> {
    await this.ensureLists();

    // Check if already dismissed
    const checkUrl = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_DISMISSED}')/items` +
      `?$select=Id&$filter=AnnouncementId eq ${announcementId} and UserId eq ${userId}&$top=1`;
    const checkResp = await this._spHttpClient.get(checkUrl, SPHttpClient.configurations.v1);
    if (checkResp.ok) {
      const checkJson = await checkResp.json();
      if ((checkJson.value || []).length > 0) return; // Already dismissed
    }

    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${LIST_DISMISSED}')/items`;
    const body = JSON.stringify({
      Title: `Dismissed ${announcementId}`,
      AnnouncementId: announcementId,
      UserId: userId
    });

    await this._spHttpClient.post(url, SPHttpClient.configurations.v1, { body });
  }

  // ─── Helpers ─────────────────────────────────────────────────────────

  /**
   * Delete a list item by ID.
   */
  private async _deleteItem(listTitle: string, itemId: number): Promise<void> {
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})`;
    const options: ISPHttpClientOptions = {
      headers: {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE'
      }
    };
    await this._spHttpClient.post(url, SPHttpClient.configurations.v1, options);
  }
}
