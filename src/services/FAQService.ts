import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

// ─── Public Interfaces ────────────────────────────────────────────────────────

export interface IFAQ {
  id: number;
  title: string;    // Question
  answer: string;
  order: number;
  category: string;
}

export interface IFAQFormData {
  title: string;
  answer: string;
  order: number;
  category: string;
}

// ─── Internal SP Response Shapes ─────────────────────────────────────────────

interface ISPFAQItem {
  ID: number;
  Title: string;
  Answer: string;
  Order0: number;
  Category: string;
}

interface ISPFAQResponse {
  value: ISPFAQItem[];
}

// ─── Service Class ────────────────────────────────────────────────────────────

export class FAQService {
  private readonly _siteUrl: string;
  private readonly _spHttpClient: SPHttpClient;
  private readonly _listName: string;

  public constructor(siteUrl: string, spHttpClient: SPHttpClient, listName: string = 'FAQs') {
    this._siteUrl = siteUrl;
    this._spHttpClient = spHttpClient;
    this._listName = listName;
  }

  /**
   * Fetch all FAQs from the SharePoint list, sorted by Order ascending.
   */
  public async getFAQs(): Promise<IFAQ[]> {
    const url =
      `${this._siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this._listName)}')/items` +
      `?$select=ID,Title,Answer,Order0,Category&$orderby=Order0 asc&$top=500`;

    const response = await this._get(url);
    if (!response.ok) {
      throw new Error(`Failed to fetch FAQs (HTTP ${response.status}).`);
    }

    const data = (await response.json()) as ISPFAQResponse;
    return data.value.map(this._mapItem);
  }

  /**
   * Add a new FAQ item to the list.
   */
  public async addFAQ(faqData: IFAQFormData): Promise<void> {
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this._listName)}')/items`;
    const body = {
      Title: faqData.title,
      Answer: faqData.answer,
      Order0: faqData.order,
      Category: faqData.category
    };

    const response = await this._spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': ''
      },
      body: JSON.stringify(body)
    });

    if (!response.ok) {
      throw new Error(`Failed to add FAQ (HTTP ${response.status}).`);
    }
  }

  /**
   * Update an existing FAQ item by its list item ID.
   */
  public async updateFAQ(id: number, faqData: IFAQFormData): Promise<void> {
    const url =
      `${this._siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this._listName)}')/items(${id})`;
    const body = {
      Title: faqData.title,
      Answer: faqData.answer,
      Order0: faqData.order,
      Category: faqData.category
    };

    const response = await this._spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify(body)
    });

    if (!response.ok) {
      throw new Error(`Failed to update FAQ (HTTP ${response.status}).`);
    }
  }

  /**
   * Delete an FAQ item by its list item ID.
   */
  public async deleteFAQ(id: number): Promise<void> {
    const url =
      `${this._siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this._listName)}')/items(${id})`;

    const response = await this._spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'odata-version': '',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE'
      },
      body: JSON.stringify({})
    });

    if (!response.ok) {
      throw new Error(`Failed to delete FAQ (HTTP ${response.status}).`);
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

  private readonly _mapItem = (item: ISPFAQItem): IFAQ => ({
    id: item.ID,
    title: item.Title || '',
    answer: item.Answer || '',
    order: item.Order0 ?? 0,
    category: item.Category || ''
  });
}
