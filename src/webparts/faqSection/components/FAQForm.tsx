import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Stack } from '@fluentui/react/lib/Stack';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';

import { IFAQ, IFAQFormData } from '../../../services/FAQService';
import importedStyles from './FAQSection.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;

interface IFAQFormProps {
  isOpen: boolean;
  initialValues: IFAQ | undefined;
  submitting: boolean;
  submitError: string | undefined;
  onSubmit: (data: IFAQFormData) => Promise<void>;
  onDismiss: () => void;
}

interface IFormState {
  title: string;
  answer: string;
  order: string;
  category: string;
}

interface IFormErrors {
  title?: string;
  answer?: string;
  order?: string;
}

const emptyState: IFormState = { title: '', answer: '', order: '0', category: '' };

type TextChangeHandler = (_ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => void;

export const FAQForm = ({
  isOpen,
  initialValues,
  submitting,
  submitError,
  onSubmit,
  onDismiss
}: IFAQFormProps): JSX.Element => {
  const [form, setForm] = React.useState<IFormState>(emptyState);
  const [errors, setErrors] = React.useState<IFormErrors>({});

  // Populate form when editing an existing item (or reset for add)
  React.useEffect(() => {
    if (isOpen) {
      if (initialValues) {
        setForm({
          title: initialValues.title,
          answer: initialValues.answer,
          order: String(initialValues.order),
          category: initialValues.category
        });
      } else {
        setForm(emptyState);
      }
      setErrors({});
    }
  }, [isOpen, initialValues]);

  const validate = (): boolean => {
    const newErrors: IFormErrors = {};
    if (!form.title.trim()) {
      newErrors.title = 'Question is required.';
    }
    if (!form.answer.trim()) {
      newErrors.answer = 'Answer is required.';
    }
    if (form.order !== '' && isNaN(Number(form.order))) {
      newErrors.order = 'Order must be a number.';
    }
    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const handleChange = (field: keyof IFormState): TextChangeHandler =>
    (_ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
      setForm(prev => ({ ...prev, [field]: newValue ?? '' }));
      setErrors(prev => ({ ...prev, [field]: undefined }));
    };

  const handleSubmit = async (): Promise<void> => {
    if (!validate()) return;

    const data: IFAQFormData = {
      title: form.title.trim(),
      answer: form.answer.trim(),
      order: form.order !== '' ? Number(form.order) : 0,
      category: form.category.trim()
    };

    await onSubmit(data);
  };

  const headerText = initialValues ? 'Edit FAQ' : 'Add FAQ';

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText={headerText}
      isLightDismiss={!submitting}
      styles={{
        header: { paddingTop: 24 },
        content: { paddingTop: 8 }
      }}
    >
      <Stack tokens={{ childrenGap: 16 }}>
        {submitError && (
          <div className={styles.formError}>
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline
            >
              {submitError}
            </MessageBar>
          </div>
        )}

        <TextField
          label="Question"
          placeholder="Enter the FAQ question"
          value={form.title}
          onChange={handleChange('title')}
          required
          errorMessage={errors.title}
          disabled={submitting}
        />

        <TextField
          label="Answer"
          placeholder="Enter the detailed answer"
          value={form.answer}
          onChange={handleChange('answer')}
          required
          multiline
          rows={6}
          autoAdjustHeight
          errorMessage={errors.answer}
          disabled={submitting}
        />

        <TextField
          label="Category"
          placeholder="e.g. General, HR, IT (leave blank for 'General')"
          value={form.category}
          onChange={handleChange('category')}
          disabled={submitting}
        />

        <TextField
          label="Order"
          placeholder="0"
          value={form.order}
          onChange={handleChange('order')}
          errorMessage={errors.order}
          disabled={submitting}
          description="Lower numbers appear first. Leave as 0 if order is not important."
          styles={{ root: { maxWidth: 200 } }}
        />

        <div className={styles.formFooter}>
          {submitting ? (
            <Spinner size={SpinnerSize.small} label="Saving…" labelPosition="right" />
          ) : (
            <>
              <PrimaryButton
                text="Save"
                onClick={handleSubmit}
                disabled={submitting}
              />
              <DefaultButton
                text="Cancel"
                onClick={onDismiss}
                disabled={submitting}
              />
            </>
          )}
        </div>
      </Stack>
    </Panel>
  );
};

export default FAQForm;
