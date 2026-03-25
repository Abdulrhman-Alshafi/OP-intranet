import * as React from 'react';
import importedStyles from './RecognitionWall.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { IKudosFormProps } from './IRecognitionWallProps';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { SPHttpClient } from '@microsoft/sp-http';

export interface IKudosFormState {
  recipientEmail: string;
  recipientName: string;
  message: string;
  category: string;
  submitting: boolean;
  error: string | null;
  searchQuery: string;
  searchResults: Array<{ displayName: string; mail: string }>;
  searching: boolean;
}

const CATEGORY_OPTIONS: IDropdownOption[] = [
  { key: 'Teamwork', text: 'Teamwork' },
  { key: 'Innovation', text: 'Innovation' },
  { key: 'Leadership', text: 'Leadership' },
  { key: 'Customer Focus', text: 'Customer Focus' },
  { key: 'Above & Beyond', text: 'Above & Beyond' }
];

/**
 * KudosForm is a slide-in panel for creating a new kudos post.
 * Includes a simple people search to find recipients.
 */
export default class KudosForm extends React.Component<IKudosFormProps, IKudosFormState> {
  private _searchTimer: number | undefined;

  constructor(props: IKudosFormProps) {
    super(props);
    this.state = {
      recipientEmail: '',
      recipientName: '',
      message: '',
      category: 'Teamwork',
      submitting: false,
      error: null,
      searchQuery: '',
      searchResults: [],
      searching: false
    };
  }

  /**
   * Search for users as the user types in the recipient field.
   * Debounced to avoid excessive API calls.
   */
  private _handleRecipientSearch = (value: string): void => {
    this.setState({ searchQuery: value, recipientEmail: '', recipientName: '' });

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

  /**
   * Call SharePoint People Search to find matching users.
   */
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

      const response = await this.props.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        body: body
      });

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

  /**
   * Select a user from the search results.
   */
  private _selectRecipient = (name: string, email: string): void => {
    this.setState({
      recipientEmail: email,
      recipientName: name,
      searchQuery: name,
      searchResults: []
    });
  };

  /**
   * Handle form submission.
   */
  private _handleSubmit = async (): Promise<void> => {
    const { recipientEmail, recipientName, message, category } = this.state;

    if (!recipientEmail || !message.trim()) {
      this.setState({ error: 'Please select a recipient and enter a message.' });
      return;
    }

    this.setState({ submitting: true, error: null });

    try {
      this.props.onSubmit(recipientEmail, recipientName, message, category);
      // Reset form
      this.setState({
        recipientEmail: '',
        recipientName: '',
        message: '',
        category: 'Teamwork',
        submitting: false,
        searchQuery: ''
      });
    } catch {
      this.setState({ submitting: false, error: 'Failed to submit kudos. Please try again.' });
    }
  };

  /**
   * Reset form on dismiss.
   */
  private _handleDismiss = (): void => {
    this.setState({
      recipientEmail: '',
      recipientName: '',
      message: '',
      category: 'Teamwork',
      error: null,
      searchQuery: '',
      searchResults: []
    });
    this.props.onDismiss();
  };

  public render(): React.ReactElement<IKudosFormProps> {
    const { isOpen } = this.props;
    const { message, category, submitting, error, searchQuery, searchResults, searching, recipientEmail } = this.state;

    return (
      <Panel
        isOpen={isOpen}
        onDismiss={this._handleDismiss}
        type={PanelType.medium}
        headerText="Give Kudos"
        closeButtonAriaLabel="Close"
      >
        <div className={styles.kudosForm}>
          {error && (
            <MessageBar messageBarType={MessageBarType.error} onDismiss={() => this.setState({ error: null })}>
              {error}
            </MessageBar>
          )}

          {/* Recipient search */}
          <div className={styles.formField}>
            <label className={styles.formLabel}>Recipient</label>
            <div className={styles.recipientSearch}>
              <input
                type="text"
                className={styles.recipientInput}
                placeholder="Search for a colleague..."
                value={searchQuery}
                onChange={(e) => this._handleRecipientSearch(e.target.value)}
                disabled={submitting}
              />
              {recipientEmail && (
                <span className={styles.selectedRecipient}>✓ {recipientEmail}</span>
              )}
              {searching && <Spinner size={SpinnerSize.xSmall} />}
              {searchResults.length > 0 && (
                <div className={styles.searchDropdown}>
                  {searchResults.map((user, idx) => (
                    <button
                      key={idx}
                      className={styles.searchResultItem}
                      onClick={() => this._selectRecipient(user.displayName, user.mail)}
                    >
                      <span className={styles.resultName}>{user.displayName}</span>
                      <span className={styles.resultEmail}>{user.mail}</span>
                    </button>
                  ))}
                </div>
              )}
            </div>
          </div>

          {/* Message */}
          <TextField
            label="Message"
            multiline
            rows={4}
            value={message}
            onChange={(_, val) => this.setState({ message: val || '' })}
            placeholder="Write your kudos message..."
            disabled={submitting}
            required
          />

          {/* Category */}
          <Dropdown
            label="Category"
            selectedKey={category}
            options={CATEGORY_OPTIONS}
            onChange={(_, opt) => { if (opt) this.setState({ category: opt.key as string }); }}
            disabled={submitting}
          />

          {/* Actions */}
          <div className={styles.formActions}>
            <PrimaryButton
              text={submitting ? 'Sending...' : 'Send Kudos'}
              onClick={() => { this._handleSubmit().catch(console.error); }}
              disabled={submitting || !recipientEmail || !message.trim()}
              iconProps={{ iconName: 'Send' }}
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
