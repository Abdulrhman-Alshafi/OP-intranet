import * as React from 'react';
import {
  Panel, PanelType
} from '@fluentui/react/lib/Panel';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';
import { Label } from '@fluentui/react/lib/Label';

const CATEGORY_OPTIONS: IDropdownOption[] = [
  { key: 'General', text: 'General' },
  { key: 'HR', text: 'HR' },
  { key: 'IT', text: 'IT' },
  { key: 'Finance', text: 'Finance' },
  { key: 'Events', text: 'Events' },
  { key: 'Policy', text: 'Policy' },
  { key: 'Urgent', text: 'Urgent' }
];

export interface IAnnouncementFormProps {
  isOpen: boolean;
  onDismiss: () => void;
  onSave: (
    title: string,
    description: string,
    category: string,
    imageUrl: string,
    isImportant: boolean
  ) => Promise<void>;
}

const AnnouncementForm: React.FC<IAnnouncementFormProps> = (props) => {
  const { isOpen, onDismiss, onSave } = props;

  const [title, setTitle] = React.useState('');
  const [description, setDescription] = React.useState('');
  const [category, setCategory] = React.useState('General');
  const [imageUrl, setImageUrl] = React.useState('');
  const [isImportant, setIsImportant] = React.useState(false);
  const [saving, setSaving] = React.useState(false);
  const [error, setError] = React.useState('');

  const resetForm = React.useCallback((): void => {
    setTitle('');
    setDescription('');
    setCategory('General');
    setImageUrl('');
    setIsImportant(false);
    setError('');
  }, []);

  const handleSave = React.useCallback(async (): Promise<void> => {
    if (!title.trim()) {
      setError('Title is required.');
      return;
    }
    if (!description.trim()) {
      setError('Description is required.');
      return;
    }

    setSaving(true);
    setError('');
    try {
      await onSave(title.trim(), description.trim(), category, imageUrl.trim(), isImportant);
      resetForm();
      onDismiss();
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to save announcement.');
    } finally {
      setSaving(false);
    }
  }, [title, description, category, imageUrl, isImportant, onSave, onDismiss, resetForm]);

  const handleDismiss = React.useCallback((): void => {
    resetForm();
    onDismiss();
  }, [onDismiss, resetForm]);

  const onRenderFooterContent = React.useCallback(
    () => (
      <Stack horizontal tokens={{ childrenGap: 10 }} styles={{ root: { padding: '16px 0' } }}>
        <PrimaryButton
          text={saving ? 'Publishing...' : 'Publish Announcement'}
          onClick={handleSave}
          disabled={saving}
          iconProps={{ iconName: 'Send' }}
          styles={{
            root: { borderRadius: 8, height: 40, paddingLeft: 20, paddingRight: 20 },
            label: { fontWeight: 600 }
          }}
        />
        <DefaultButton
          text="Cancel"
          onClick={handleDismiss}
          disabled={saving}
          styles={{ root: { borderRadius: 8, height: 40 } }}
        />
      </Stack>
    ),
    [saving, handleSave, handleDismiss]
  );

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={handleDismiss}
      type={PanelType.medium}
      headerText="New Announcement"
      closeButtonAriaLabel="Close"
      onRenderFooterContent={onRenderFooterContent}
      isFooterAtBottom={true}
      styles={{
        main: { borderTopLeftRadius: 12, borderBottomLeftRadius: 12 },
        headerText: { fontSize: 22, fontWeight: 700 }
      }}
    >
      <Stack tokens={{ childrenGap: 20 }} styles={{ root: { paddingTop: 8 } }}>
        {error && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
            {error}
          </MessageBar>
        )}

        <TextField
          label="Title"
          required
          value={title}
          onChange={(_, val) => setTitle(val || '')}
          placeholder="Enter announcement title"
          maxLength={255}
          styles={{
            fieldGroup: { borderRadius: 8, border: '1px solid #e1dfdd' },
            field: { padding: '8px 12px' }
          }}
        />

        <TextField
          label="Description"
          required
          multiline
          rows={5}
          value={description}
          onChange={(_, val) => setDescription(val || '')}
          placeholder="Enter announcement details..."
          styles={{
            fieldGroup: { borderRadius: 8, border: '1px solid #e1dfdd' },
            field: { padding: '8px 12px' }
          }}
        />

        <Dropdown
          label="Category"
          selectedKey={category}
          options={CATEGORY_OPTIONS}
          onChange={(_, opt) => { if (opt) setCategory(opt.key as string); }}
          styles={{
            title: { borderRadius: 8, border: '1px solid #e1dfdd', height: 38 },
            dropdown: { borderRadius: 8 }
          }}
        />

        <TextField
          label="Image URL (optional)"
          placeholder="Paste high quality image URL here..."
          value={imageUrl}
          onChange={(_, val) => setImageUrl(val || '')}
          styles={{
            fieldGroup: { borderRadius: 8, border: '1px solid #e1dfdd' },
            field: { padding: '8px 12px' }
          }}
        />

        <Toggle
          label="Mark as Important"
          checked={isImportant}
          onChange={(_, checked) => setIsImportant(!!checked)}
          onText="Yes — highlighted to all users"
          offText="No"
          styles={{ root: { marginBottom: 0 } }}
        />
      </Stack>
    </Panel>
  );
};

export default AnnouncementForm;
