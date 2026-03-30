import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';

// ── Public Data Shapes ────────────────────────────────────────────────────────

export interface ITicket {
  /** e.g. "R-000042" */
  ref: string;
  /** iTop internal numeric id */
  id: string;
  title: string;
  /** raw iTop status string: "new" | "assigned" | "resolved" | "closed" | "pending" */
  status: string;
  /** "1" (critical) … "4" (low) — mapped to text by the UI */
  urgency: string;
  urgencyLabel: string;
  service: string;
  lastUpdate: string;
  startDate: string;
  /** Deep-link into the iTop portal for this ticket */
  deepLink: string;
}

export interface IPagedTickets {
  items: ITicket[];
  totalCount: number;
  hasMore: boolean;
}

export interface ITicketServiceItem {
  id: string;
  name: string;
}

export interface ICreateTicketData {
  title: string;
  description: string;
  serviceId: string;
  urgency: string;
  callerEmail: string;
}

// ── Service Class ─────────────────────────────────────────────────────────────

export class ITopTicketService {
  private readonly _baseUrl: string;
  private readonly _functionKey: string;
  private readonly _httpClient: HttpClient;

  private _servicesCache: ITicketServiceItem[] | undefined;

  constructor(
    functionBaseUrl: string,
    functionKey: string,
    httpClient: HttpClient
  ) {
    this._baseUrl = functionBaseUrl.replace(/\/$/, '');
    this._functionKey = functionKey;
    this._httpClient = httpClient;
  }

  // ── Internals ───────────────────────────────────────────────────────────────

  private _buildUrl(
    path: string,
    params: Record<string, string> = {}
  ): string {
    const url = new URL(`${this._baseUrl}${path}`);
    const paramKeys = Object.keys(params);
    for (let i = 0; i < paramKeys.length; i++) {
      const k = paramKeys[i];
      const v = params[k];
      if (v !== undefined && v !== null && v !== '') {
        url.searchParams.set(k, v);
      }
    }
    return url.toString();
  }

  /** Common headers — function key via x-functions-key header (not query param). */
  private _authHeaders(): Record<string, string> {
    const h: Record<string, string> = {};
    if (this._functionKey) {
      h['x-functions-key'] = this._functionKey;
    }
    return h;
  }

  private async _get<T>(
    path: string,
    params: Record<string, string> = {}
  ): Promise<T> {
    const url = this._buildUrl(path, params);
    const response: HttpClientResponse = await this._httpClient.get(
      url,
      HttpClient.configurations.v1,
      { headers: this._authHeaders() }
    );
    if (!response.ok) {
      const text = await response.text().catch(() => String(response.status));
      throw new Error(`[iTop proxy] GET ${path} → HTTP ${response.status}: ${text}`);
    }
    return response.json() as Promise<T>;
  }

  private async _post<T>(path: string, body: unknown): Promise<T> {
    const url = this._buildUrl(path);
    const headers = this._authHeaders();
    headers['Content-Type'] = 'application/json';
    const response: HttpClientResponse = await this._httpClient.post(
      url,
      HttpClient.configurations.v1,
      {
        body: JSON.stringify(body),
        headers
      }
    );
    if (!response.ok) {
      const text = await response.text().catch(() => String(response.status));
      throw new Error(`[iTop proxy] POST ${path} → HTTP ${response.status}: ${text}`);
    }
    return response.json() as Promise<T>;
  }

  // ── Public API ──────────────────────────────────────────────────────────────

  /** Returns tickets where caller_id.email = email */
  public async getMyTickets(email: string): Promise<ITicket[]> {
    return this._get<ITicket[]>('/api/itop/tickets', { email });
  }

  /** Returns all tickets (IT Admin view) with pagination */
  public async getAllTickets(
    page: number = 1,
    limit: number = 50
  ): Promise<IPagedTickets> {
    return this._get<IPagedTickets>('/api/itop/tickets', {
      all: 'true',
      page: String(page),
      limit: String(limit)
    });
  }

  /** Returns the list of iTop services for the New Ticket dropdown. Cached. */
  public async getServices(): Promise<ITicketServiceItem[]> {
    if (this._servicesCache) return this._servicesCache;
    const services = await this._get<ITicketServiceItem[]>('/api/itop/services');
    this._servicesCache = services;
    return services;
  }

  /** Creates a new UserRequest ticket */
  public async createTicket(data: ICreateTicketData): Promise<ITicket> {
    return this._post<ITicket>('/api/itop/tickets', data);
  }
}
