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

/**
 * Panel for admins to select Employee of the Month via people search.
 */
const EmployeeOfMonthForm = (props: IEmployeeOfMonthFormProps): React.ReactElement => {
  const { isOpen, onDismiss, onSubmit, siteUrl, spHttpClient } = props;
  const [searchQuery, setSearchQuery] = React.useState<string>('');
  const [searchResults, setSearchResults] = React.useState<Array<{ displayName: string; mail: string }>>([]);
  const [searching, setSearching] = React.useState<boolean>(false);
  const [selectedEmail, setSelectedEmail] = React.useState<string>('');
  const [selectedName, setSelectedName] = React.useState<string>('');
  const [submitting, setSubmitting] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | null>(null);
  const searchTimerRef = React.useRef<number | undefined>(undefined);

  const resetForm = (): void => {
    setSearchQuery('');
    setSelectedEmail('');
    setSelectedName('');
    setSearchResults([]);
    setError(null);
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

      const response = await spHttpClient.post(url, SPHttpClient.configurations.v1, { body });

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

  const handleSearch = (value: string): void => {
    setSearchQuery(value);
    setSelectedEmail('');
    setSelectedName('');

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

  const selectPerson = (name: string, email: string): void => {
    setSelectedEmail(email);
    setSelectedName(name);
    setSearchQuery(name);
    setSearchResults([]);
  };

  const handleSubmit = async (): Promise<void> => {
    if (!selectedEmail) {
      setError('Please select an employee.');
      return;
    }

    setSubmitting(true);
    setError(null);

    try {
      await onSubmit(selectedEmail, selectedName);
      setSearchQuery('');
      setSelectedEmail('');
      setSelectedName('');
      setSubmitting(false);
    } catch {
      setSubmitting(false);
      setError('Failed to set Employee of the Month.');
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
      headerText="Select Employee of the Month"
      closeButtonAriaLabel="Close"
    >
      <div className={styles.kudosForm}>
        {error && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(null)}>
            {error}
          </MessageBar>
        )}

        <div className={styles.formField}>
          <label className={styles.formLabel}>Employee</label>
          <div className={styles.recipientSearch}>
            <input
              type="text"
              className={styles.recipientInput}
              placeholder="Search for an employee..."
              value={searchQuery}
              onChange={(e) => handleSearch(e.target.value)}
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
                    onClick={() => selectPerson(user.displayName, user.mail)}
                  >
                    <span className={styles.resultName}>{user.displayName}</span>
                    <span className={styles.resultEmail}>{user.mail}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
        </div>

        <div className={styles.formActions}>
          <PrimaryButton
            text={submitting ? 'Setting...' : 'Set Employee of the Month'}
            onClick={() => {
              handleSubmit().catch(() => {
                setSubmitting(false);
                setError('Failed to set Employee of the Month.');
              });
            }}
            disabled={submitting || !selectedEmail}
            iconProps={{ iconName: 'FavoriteStar' }}
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

export default EmployeeOfMonthForm;
