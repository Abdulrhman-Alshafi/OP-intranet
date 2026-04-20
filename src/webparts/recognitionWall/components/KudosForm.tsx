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
const KudosForm = (props: IKudosFormProps): React.ReactElement => {
  const { isOpen, onDismiss, onSubmit, siteUrl, spHttpClient } = props;
  const [recipientEmail, setRecipientEmail] = React.useState<string>('');
  const [recipientName, setRecipientName] = React.useState<string>('');
  const [message, setMessage] = React.useState<string>('');
  const [category, setCategory] = React.useState<string>('Teamwork');
  const [submitting, setSubmitting] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | null>(null);
  const [searchQuery, setSearchQuery] = React.useState<string>('');
  const [searchResults, setSearchResults] = React.useState<Array<{ displayName: string; mail: string }>>([]);
  const [searching, setSearching] = React.useState<boolean>(false);
  const searchTimerRef = React.useRef<number | undefined>(undefined);

  const resetForm = (): void => {
    setRecipientEmail('');
    setRecipientName('');
    setMessage('');
    setCategory('Teamwork');
    setError(null);
    setSearchQuery('');
    setSearchResults([]);
  };

  React.useEffect(() => {
    return () => {
      if (searchTimerRef.current) {
        window.clearTimeout(searchTimerRef.current);
      }
    };
  }, []);

  const searchUsers = async (query: string): Promise<void> => {
    setSearching(true);

    try {
      const url = `${siteUrl}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.ClientPeoplePickerSearchUser`;
      const body = JSON.stringify({
        queryParams: {
          QueryString: query,
          MaximumEntitySuggestions: 5,
          AllowEmailAddresses: true,
          PrincipalType: 1,
          PrincipalSource: 15
        }
      });

      const response = await spHttpClient.post(url, SPHttpClient.configurations.v1, {
        body
      });

      if (response.ok) {
        const json = await response.json();
        const resultStr = json.value || json.ClientPeoplePickerSearchUser || '[]';
        const parsed = typeof resultStr === 'string' ? JSON.parse(resultStr) : resultStr;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const results = (parsed as any[]).map((person) => ({
          displayName: person.DisplayText || person.Key || '',
          mail: person.EntityData?.Email || person.Key || ''
        })).filter((result) => result.mail);
        setSearchResults(results);
        setSearching(false);
      } else {
        setSearchResults([]);
        setSearching(false);
      }
    } catch {
      setSearchResults([]);
      setSearching(false);
    }
  };

  const handleRecipientSearch = (value: string): void => {
    setSearchQuery(value);
    setRecipientEmail('');
    setRecipientName('');

    if (searchTimerRef.current) {
      window.clearTimeout(searchTimerRef.current);
    }

    if (value.length < 2) {
      setSearchResults([]);
      setSearching(false);
      return;
    }

    searchTimerRef.current = window.setTimeout(() => {
      searchUsers(value).catch(() => {
        setSearchResults([]);
        setSearching(false);
      });
    }, 400);
  };

  const selectRecipient = (name: string, email: string): void => {
    setRecipientEmail(email);
    setRecipientName(name);
    setSearchQuery(name);
    setSearchResults([]);
  };

  const handleSubmit = async (): Promise<void> => {
    if (!recipientEmail || !message.trim()) {
      setError('Please select a recipient and enter a message.');
      return;
    }

    setSubmitting(true);
    setError(null);

    try {
      onSubmit(recipientEmail, recipientName, message, category);
      setRecipientEmail('');
      setRecipientName('');
      setMessage('');
      setCategory('Teamwork');
      setSubmitting(false);
      setSearchQuery('');
    } catch {
      setSubmitting(false);
      setError('Failed to submit kudos. Please try again.');
    }
  };

  const handleDismiss = (): void => {
    resetForm();
    onDismiss();
  };

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={handleDismiss}
      type={PanelType.medium}
      headerText="Give Kudos"
      closeButtonAriaLabel="Close"
    >
      <div className={styles.kudosForm}>
        {error && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(null)}>
            {error}
          </MessageBar>
        )}

        <div className={styles.formField}>
          <label className={styles.formLabel}>Recipient</label>
          <div className={styles.recipientSearch}>
            <input
              type="text"
              className={styles.recipientInput}
              placeholder="Search for a colleague..."
              value={searchQuery}
              onChange={(e) => handleRecipientSearch(e.target.value)}
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
                    onClick={() => selectRecipient(user.displayName, user.mail)}
                  >
                    <span className={styles.resultName}>{user.displayName}</span>
                    <span className={styles.resultEmail}>{user.mail}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>

        <TextField
          label="Message"
          multiline
          rows={4}
          value={message}
          onChange={(_, val) => setMessage(val || '')}
          placeholder="Write your kudos message..."
          disabled={submitting}
          required
        />

        <Dropdown
          label="Category"
          selectedKey={category}
          options={CATEGORY_OPTIONS}
          onChange={(_, opt) => { if (opt) setCategory(opt.key as string); }}
          disabled={submitting}
        />

        <div className={styles.formActions}>
          <PrimaryButton
            text={submitting ? 'Sending...' : 'Send Kudos'}
            onClick={() => {
              handleSubmit().catch(() => {
                setSubmitting(false);
                setError('Failed to submit kudos. Please try again.');
              });
            }}
            disabled={submitting || !recipientEmail || !message.trim()}
            iconProps={{ iconName: 'Send' }}
          />
          <DefaultButton
            text="Cancel"
            onClick={handleDismiss}
            disabled={submitting}
          />
          {submitting && <Spinner size={SpinnerSize.small} />}
        </div>
      </div>
    </Panel>
  );
};

export default KudosForm;
