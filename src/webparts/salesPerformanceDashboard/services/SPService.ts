import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ISalesData } from '../models/ISalesData';

export class SPService {
  public static async getSalesData(context: WebPartContext, listId: string): Promise<ISalesData[]> {
    if (!listId) return [];
    
    try {
      const endpoint = `${context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${listId}')/items?$select=Id,Title,Salesperson,Region,Product,Target,Achieved,Date`;
      const response: SPHttpClientResponse = await context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
      
      if (response.ok) {
        const json = await response.json();
        return json.value as ISalesData[];
      }
      return [];
    } catch (error) {
      console.error("Error fetching sales data: ", error);
      return [];
    }
  }
}
