import * as React from 'react';
import importedStyles from './RecognitionWall.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { SPHttpClient } from '@microsoft/sp-http';

export interface IEmployeeOfMonthFormProps {
  isOpen: boolean;
  onDismiss: () => void;
  onSubmit: (email: string, name: string) => Promise<void>;
  siteUrl: string;
  spHttpClient: SPHttpClient;
}

export interface IEmployeeOfMonthFormState {
  searchQuery: string;
  searchResults: Array<{ displayName: string; mail: string }>;
  searching: boolean;
  selectedEmail: string;
  selectedName: string;
  submitting: boolean;
  error: string | null;
}

/**
 * Panel for admins to select Employee of the Month via people search.
 */
export default class EmployeeOfMonthForm extends React.Component<IEmployeeOfMonthFormProps, IEmployeeOfMonthFormState> {
  private _searchTimer: number | undefined;

  constructor(props: IEmployeeOfMonthFormProps) {
    super(props);
    this.state = {
      searchQuery: '',
      searchResults: [],
      searching: false,
      selectedEmail: '',
      selectedName: '',
      submitting: false,
      error: null
    };
  }

  private _handleSearch = (value: string): void => {
    this.setState({ searchQuery: value, selectedEmail: '', selectedName: '' });

    if (this._searchTimer) {
      window.clearTimeout(this._searchTimer);
    }

    if (value.length < 2) {
      this.setState({ searchResults: [], searching: false });
      return;
    }

    this._searchTimer = window.setTimeout(() => {
      this._searchUsers(value).catch(console.error);
    }, 400);
  };

  private async _searchUsers(query: string): Promise<void> {
    this.setState({ searching: true });

    try {
      const url = `${this.props.siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser`;
      const body = JSON.stringify({
        queryParams: {
          QueryString: query,
          MaximumEntitySuggestions: 5,
          AllowEmailAddresses: true,
          PrincipalType: 1,
          PrincipalSource: 15
        }
      });

      const response = await this.props.spHttpClient.post(url, SPHttpClient.configurations.v1, { body });

      if (response.ok) {
        const json = await response.json();
        const resultStr = json.value || json.ClientPeoplePickerSearchUser || '[]';
        const parsed = typeof resultStr === 'string' ? JSON.parse(resultStr) : resultStr;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const results = (parsed as any[]).map(p => ({
          displayName: p.DisplayText || p.Key || '',
          mail: p.EntityData?.Email || p.Key || ''
        })).filter(r => r.mail);
        this.setState({ searchResults: results, searching: false });
      } else {
        this.setState({ searchResults: [], searching: false });
      }
    } catch {
      this.setState({ searchResults: [], searching: false });
    }
  }

  private _selectPerson = (name: string, email: string): void => {
    this.setState({
      selectedEmail: email,
      selectedName: name,
      searchQuery: name,
      searchResults: []
    });
  };

  private _handleSubmit = async (): Promise<void> => {
    const { selectedEmail, selectedName } = this.state;
    if (!selectedEmail) {
      this.setState({ error: 'Please select an employee.' });
      return;
    }

    this.setState({ submitting: true, error: null });

    try {
      await this.props.onSubmit(selectedEmail, selectedName);
      this.setState({
        searchQuery: '',
        selectedEmail: '',
        selectedName: '',
        submitting: false
      });
    } catch {
      this.setState({ submitting: false, error: 'Failed to set Employee of the Month.' });
    }
  };

  private _handleDismiss = (): void => {
    this.setState({
      searchQuery: '',
      selectedEmail: '',
      selectedName: '',
      searchResults: [],
      error: null
    });
    this.props.onDismiss();
  };

  public render(): React.ReactElement<IEmployeeOfMonthFormProps> {
    const { isOpen } = this.props;
    const { searchQuery, searchResults, searching, selectedEmail, submitting, error } = this.state;

    return (
      <Panel
        isOpen={isOpen}
        onDismiss={this._handleDismiss}
        type={PanelType.medium}
        headerText="Select Employee of the Month"
        closeButtonAriaLabel="Close"
      >
        <div className={styles.kudosForm}>
          {error && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ error: null })}>
              {error}
            </MessageBar>
          )}

          {/* Employee search */}
          <div className={styles.formField}>
            <label className={styles.formLabel}>Employee</label>
            <div className={styles.recipientSearch}>
              <input
                type="text"
                className={styles.recipientInput}
                placeholder="Search for an employee..."
                value={searchQuery}
                onChange={(e) => this._handleSearch(e.target.value)}
                disabled={submitting}
              />
              {selectedEmail && (
                <span className={styles.selectedRecipient}>✓ {selectedEmail}</span>
              )}
              {searching && <Spinner size={SpinnerSize.xSmall} />}
              {searchResults.length > 0 && (
                <div className={styles.searchDropdown}>
                  {searchResults.map((user, idx) => (
                    <button
                      key={idx}
                      className={styles.searchResultItem}
                      onClick={() => this._selectPerson(user.displayName, user.mail)}
                    >
                      <span className={styles.resultName}>{user.displayName}</span>
                      <span className={styles.resultEmail}>{user.mail}</span>
                    </button>
                  ))}
                </div>
              )}
            </div>
          </div>

          {/* Actions */}
          <div className={styles.formActions}>
            <PrimaryButton
              text={submitting ? 'Setting...' : 'Set Employee of the Month'}
              onClick={() => { this._handleSubmit().catch(console.error); }}
              disabled={submitting || !selectedEmail}
              iconProps={{ iconName: 'FavoriteStar' }}
            />
            <DefaultButton
              text="Cancel"
              onClick={this._handleDismiss}
              disabled={submitting}
            />
            {submitting && <Spinner size={SpinnerSize.small} />}
          </div>
        </div>
      </Panel>
    );
  }
}
