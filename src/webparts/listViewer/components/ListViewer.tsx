import * as React from 'react';
import styles from './ListViewer.module.scss';
import type { IListViewerProps } from './IListViewerProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { escape } from '@microsoft/sp-lodash-subset';

export interface IListViewerState {
  items: any[];
  loading: boolean;
  error: string | null;
}

export default class ListViewer extends React.Component<IListViewerProps, IListViewerState> {
  constructor(props: IListViewerProps) {
    super(props);
    this.state = {
      items: [],
      loading: false,
      error: null
    };
  }

  public componentDidMount(): void {
    if (this.props.listId) {
      this._fetchItems().catch(console.error);
    }
  }

  public componentDidUpdate(prevProps: IListViewerProps): void {
    if (prevProps.listId !== this.props.listId || prevProps.maxItems !== this.props.maxItems || prevProps.siteUrl !== this.props.siteUrl) {
      if (this.props.listId) {
        this._fetchItems().catch(console.error);
      } else {
        this.setState({ items: [], error: null });
      }
    }
  }

  private _fetchItems = async (): Promise<void> => {
    this.setState({ loading: true, error: null });
    
    try {
      const endpoint = `${this.props.siteUrl}/_api/web/lists(guid'${this.props.listId}')/items?$select=Id,Title,Created,Author/Title&$expand=Author&$top=${this.props.maxItems || 10}&$orderby=Created desc`;
      const response: SPHttpClientResponse = await this.props.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
      
      if (response.ok) {
        const json = await response.json();
        this.setState({ items: json.value || [], loading: false });
      } else {
        const errJson = await response.json().catch(() => ({}));
        this.setState({ error: errJson.error?.message?.value || 'Failed to fetch items', loading: false });
      }
    } catch (err: any) {
      this.setState({ error: err.message, loading: false });
    }
  };

  private _formatDate(dateString: string): string {
    if (!dateString) return '';
    return new Date(dateString).toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' });
  }

  public render(): React.ReactElement<IListViewerProps> {
    const { title, viewStyle, hasTeamsContext, listId } = this.props;
    const { items, loading, error } = this.state;

    return (
      <section className={`${styles.container} ${hasTeamsContext ? styles.teams : ''}`}>
        {title && <h2 className={styles.header}>{escape(title)}</h2>}
        
        {!listId && (
          <div className={styles.emptyState}>
            Please select a list from the property pane to view its items.
          </div>
        )}

        {loading && <div className={styles.emptyState}>Loading items...</div>}
        
        {error && <div className={styles.emptyState} style={{ color: 'red' }}>Error: {error}</div>}

        {!loading && !error && listId && items.length === 0 && (
          <div className={styles.emptyState}>No items found in the selected list.</div>
        )}

        {!loading && !error && items.length > 0 && viewStyle === 'table' && (
          <div className={styles.tableContainer}>
            <table className={styles.table}>
              <thead>
                <tr>
                  <th>Title</th>
                  <th>Created By</th>
                  <th>Creation Date</th>
                </tr>
              </thead>
              <tbody>
                {items.map((item, index) => (
                  <tr key={index}>
                    <td style={{ fontWeight: 600 }}>{item.Title || `Item ${item.Id}`}</td>
                    <td>{item.Author?.Title || 'Unknown'}</td>
                    <td>{this._formatDate(item.Created)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        {!loading && !error && items.length > 0 && viewStyle !== 'table' && (
          <div className={viewStyle === 'list' ? styles.list : styles.grid}>
            {items.map((item, index) => (
              <div key={index} className={styles.card}>
                <h3 className={styles.cardTitle}>{item.Title || `Item ${item.Id}`}</h3>
                <div className={styles.cardMeta}>
                  <svg width="12" height="12" viewBox="0 0 2048 2048" fill="currentColor"><path d="M1024 0q141 0 272 36t244 104 207 160 161 207 103 245 37 272q0 141-36 272t-104 244-160 207-207 161-245 103-272 37q-141 0-272-36t-244-104-207-160-161-207-103-245-37-272q0-141 36-272t104-244 160-207 207-161 245-103 272-37zm0 1920q123 0 237-32t214-90 182-141 140-181 91-214 32-238q0-123-32-237t-90-214-141-182-181-140-214-91-238-32q-123 0-237 32t-214 90-182 141-140 181-91 214-32 238q0 123 32 237t90 214 141 182 181 140 214 91 238 32zm128-1024h512v128h-640v-640h128v512z"/></svg>
                  {this._formatDate(item.Created)}
                </div>
                {item.Author && item.Author.Title && (
                  <div className={styles.cardMeta}>
                    <svg width="12" height="12" viewBox="0 0 2048 2048" fill="currentColor"><path d="M1024 1024q141 0 272-36t244-104 207-160 161-207 103-245 37-272q0-141-36-272t-104-244-160-207-207-161-245-103-272-37q-141 0-272 36t-244 104-207 160-161 207-103 245-37 272q0 141 36 272t104 244 160 207 207 161 245 103 272 37zm0-896q123 0 237 32t214 90 182 141 140 181 91 214 32 238q0 123-32 237t-90 214-141 182-181 140-214 91-238 32q-123 0-237-32t-214-90-182-141-140-181-91-214-32-238q0-123 32-237t90-214 141-182 181-140 214-91 238-32zm0 1920q-137 0-268-36t-246-103-207-162-161-207-103-246-37-268h128q0 119 32 231t90 209 141 178 181 138 214 88 238 31q119 0 231-32t209-90 178-141 138-181 88-214 31-238h128q0 137-36 268t-103 246-162 207-207 161-246 103-268 37z"/></svg>
                    {item.Author.Title}
                  </div>
                )}
              </div>
            ))}
          </div>
        )}
      </section>
    );
  }
}
