import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IKBArticle {
  Id: number;
  Title: string;
  Description: string;
  Category: string;
  Content: string;
  Created: string;
}

interface ISPListItem {
  Id: number;
  Title: string;
  Description: string;
  Category: string;
  Content: string;
  Created: string;
}

interface ISPListResponse {
  value: ISPListItem[];
}

export class KnowledgeBaseService {
  public static async getArticles(
    siteUrl: string,
    listName: string,
    spHttpClient: SPHttpClient
  ): Promise<IKBArticle[]> {
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listName)}')/items?$select=Id,Title,Description,Category,Content,Created&$orderby=Created desc&$top=500`;

    const response: SPHttpClientResponse = await spHttpClient.get(
      endpoint,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to fetch articles: ${response.status} – ${errorText}`);
    }

    const data: ISPListResponse = await response.json();

    return data.value.map((item) => ({
      Id: item.Id,
      Title: item.Title || '',
      Description: item.Description || '',
      Category: item.Category || 'General',
      Content: item.Content || '',
      Created: item.Created || ''
    }));
  }

  public static getDistinctCategories(articles: IKBArticle[]): string[] {
    const categorySet = new Set<string>();
    articles.forEach((a) => {
      if (a.Category) categorySet.add(a.Category);
    });
    return Array.from(categorySet).sort();
  }
}
