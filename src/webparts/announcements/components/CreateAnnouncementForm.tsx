import * as React from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import importedStyles from './Announcements.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;

interface ICreateAnnouncementFormProps {
  onSubmit: (title: string, content: string) => Promise<void>;
  loading?: boolean;
}

const CreateAnnouncementForm: React.FC<ICreateAnnouncementFormProps> = ({ onSubmit, loading }) => {
  const [title, setTitle] = React.useState<string>('');
  const [content, setContent] = React.useState<string>('');
  const [submitting, setSubmitting] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>('');

  const handleSubmit = async (): Promise<void> => {
    if (!title.trim() || !content.trim()) {
      setError('Please fill in both title and content');
      return;
    }

    setSubmitting(true);
    setError('');

    try {
      await onSubmit(title, content);
      setTitle('');
      setContent('');
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to create announcement');
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div className={styles.createAnnouncementForm}>
      {error && <div className={styles.errorMessage}>{error}</div>}

      <TextField
        label="Announcement Title"
        placeholder="What's new?"
        value={title}
        onChange={(_ev, newValue) => setTitle(newValue || '')}
        disabled={submitting || loading}
        multiline={false}
      />

      <TextField
        label="Content"
        placeholder="Share your message here..."
        value={content}
        onChange={(_ev, newValue) => setContent(newValue || '')}
        disabled={submitting || loading}
        multiline
        rows={4}
        styles={{ root: { marginTop: 12, marginBottom: 12 } }}
      />

      <div style={{ display: 'flex', gap: 8 }}>
        <PrimaryButton
          text="Post"
          onClick={handleSubmit}
          disabled={submitting || loading || !title.trim() || !content.trim()}
          iconProps={{ iconName: 'Send' }}
        />
        {(submitting || loading) && <span style={{ marginTop: 8 }}>Posting...</span>}
      </div>
    </div>
  );
};

export default CreateAnnouncementForm;
