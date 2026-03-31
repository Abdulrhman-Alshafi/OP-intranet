import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IHelpDeskDevice {
  Id: number;
  Title: string; // Device Name/Model
  DeviceType: string;
  SerialNumber?: string;
  IsRequestable: boolean;
  Status: string;
}

export interface IHelpDeskDeviceRequest {
  Id: number;
  Title: string; // Reason
  DeviceId?: number;
  RequesterId?: number;
  RequesterName?: string;
  Status: string; // Pending, Approved, Rejected
  DeviceName?: string; // Optional field mapping
}

export class HelpDeskDeviceService {
  private readonly _siteUrl: string;
  private readonly _spHttpClient: SPHttpClient;

  public constructor(siteUrl: string, spHttpClient: SPHttpClient) {
    this._siteUrl = siteUrl;
    this._spHttpClient = spHttpClient;
  }

  public async isUserInGroup(groupName: string): Promise<boolean> {
    if (!groupName) return false;
    const encoded = encodeURIComponent(groupName);
    const url = `${this._siteUrl}/_api/web/currentuser/groups?$select=Title`;
    try {
      const response = await this._get(url);
      if (!response.ok) return false;
      const data = (await response.json()) as { value: { Title: string }[] };
      return (data.value || []).some((g) => g.Title === groupName);
    } catch {
      return false;
    }
  }

  public async isCurrentUserSiteAdmin(): Promise<boolean> {
    const url = `${this._siteUrl}/_api/web/currentuser?$select=IsSiteAdmin`;
    try {
      const response = await this._get(url);
      if (!response.ok) return false;
      const data = (await response.json()) as { IsSiteAdmin?: boolean };
      return data.IsSiteAdmin === true;
    } catch {
      return false;
    }
  }

  public async getDevices(listName: string): Promise<IHelpDeskDevice[]> {
    const encoded = encodeURIComponent(listName);
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${encoded}')/items?$select=ID,Title,DeviceType,SerialNumber,IsRequestable,Status&$top=5000`;
    const response = await this._get(url);
    if (!response.ok) return [];
    const data = await response.json();
    return (data.value || []).map((item: any) => ({
      Id: item.ID,
      Title: item.Title ?? '',
      DeviceType: item.DeviceType ?? '',
      SerialNumber: item.SerialNumber ?? '',
      IsRequestable: item.IsRequestable === true,
      Status: item.Status ?? 'Available'
    }));
  }

  public async addDevice(listName: string, device: Omit<IHelpDeskDevice, 'Id'>): Promise<void> {
    const encoded = encodeURIComponent(listName);
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${encoded}')/items`;
    const response = await this._post(url, {
      Title: device.Title,
      DeviceType: device.DeviceType,
      SerialNumber: device.SerialNumber,
      IsRequestable: device.IsRequestable,
      Status: device.Status
    });
    if (!response.ok) {
      throw new Error(`Failed to add device: HTTP ${response.status}`);
    }
  }

  public async updateDeviceIsRequestable(listName: string, deviceId: number, isRequestable: boolean): Promise<void> {
    const encoded = encodeURIComponent(listName);
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${encoded}')/items(${deviceId})`;
    const response = await this._spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify({ IsRequestable: isRequestable })
    });
    if (!response.ok) {
      throw new Error(`Failed to update device: HTTP ${response.status}`);
    }
  }

  public async getRequests(listName: string): Promise<IHelpDeskDeviceRequest[]> {
    const encoded = encodeURIComponent(listName);
    // Assuming Requester is a Person field. We expand it.
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${encoded}')/items?$select=ID,Title,DeviceId,RequesterId,Requester/Title,Status&$expand=Requester&$top=5000`;
    const response = await this._get(url);
    if (!response.ok) return [];
    let data;
    try { data = await response.json(); } catch(e) { return []; }
    return (data.value || []).map((item: any) => ({
      Id: item.ID,
      Title: item.Title ?? '',
      DeviceId: item.DeviceId,
      RequesterId: item.RequesterId,
      RequesterName: item.Requester ? item.Requester.Title : 'Unknown',
      Status: item.Status ?? 'Pending'
    }));
  }

  public async updateRequestStatus(listName: string, requestId: number, status: string): Promise<void> {
    const encoded = encodeURIComponent(listName);
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${encoded}')/items(${requestId})`;
    const response = await this._spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify({ Status: status })
    });
    if (!response.ok) {
      throw new Error(`Failed to update request: HTTP ${response.status}`);
    }
  }

  // ── Added Methods for User Requests ─────────────────────────────────────

  public async getCurrentUserId(): Promise<number> {
    const url = `${this._siteUrl}/_api/web/currentuser?$select=Id`;
    const response = await this._get(url);
    if (!response.ok) throw new Error('Could not get current user');
    const data = await response.json();
    return data.Id;
  }

  public async getMyRequests(listName: string, currentUserId: number): Promise<IHelpDeskDeviceRequest[]> {
    const encoded = encodeURIComponent(listName);
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${encoded}')/items?$select=ID,Title,DeviceId,RequesterId,Requester/Title,Status&$expand=Requester&$filter=RequesterId eq ${currentUserId}&$top=5000`;
    const response = await this._get(url);
    if (!response.ok) return [];
    let data;
    try { data = await response.json(); } catch(e) { return []; }
    return (data.value || []).map((item: any) => ({
      Id: item.ID,
      Title: item.Title ?? '',
      DeviceId: item.DeviceId,
      RequesterId: item.RequesterId,
      RequesterName: item.Requester ? item.Requester.Title : 'Unknown',
      Status: item.Status ?? 'Pending'
    }));
  }

  public async addRequest(listName: string, request: { DeviceId: number, RequesterId: number, Title: string }): Promise<void> {
    const encoded = encodeURIComponent(listName);
    const url = `${this._siteUrl}/_api/web/lists/getbytitle('${encoded}')/items`;
    const response = await this._post(url, {
      Title: request.Title,
      DeviceId: request.DeviceId,
      RequesterId: request.RequesterId,
      Status: 'Pending'
    });
    if (!response.ok) {
      throw new Error(`Failed to add request: HTTP ${response.status}`);
    }
  }

  private async _get(url: string): Promise<SPHttpClientResponse> {
    return this._spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata=nometadata', 'odata-version': '' }
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
}
