import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { Text } from '@fluentui/react/lib/Text';

import { IFAQSectionProps } from './IFAQSectionProps';
import { IFAQ, IFAQFormData } from '../../../services/FAQService';
import { FAQList } from './FAQList';
import { FAQForm } from './FAQForm';
import importedStyles from './FAQSection.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;

export const FAQSection = ({ service }: IFAQSectionProps): JSX.Element => {
  const [faqs, setFaqs] = React.useState<IFAQ[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [canEdit, setCanEdit] = React.useState(false);
  const [expandedIds, setExpandedIds] = React.useState<Set<number>>(new Set());
  const [isPanelOpen, setIsPanelOpen] = React.useState(false);
  const [editingFaq, setEditingFaq] = React.useState<IFAQ | undefined>(undefined);
  const [submitting, setSubmitting] = React.useState(false);
  const [submitError, setSubmitError] = React.useState<string | undefined>(undefined);

  const loadData = React.useCallback(async (): Promise<void> => {
    setLoading(true);
    setError(undefined);
    try {
      const [fetchedFaqs, isAdmin] = await Promise.all([
        service.getFAQs(),
        service.isCurrentUserSiteAdmin()
      ]);
      setFaqs(fetchedFaqs);
      setCanEdit(isAdmin);
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : 'An unexpected error occurred while loading FAQs.'
      );
      setFaqs([]);
    } finally {
      setLoading(false);
    }
  }, [service]);

  React.useEffect(() => {
    loadData().catch(() => { /* handled inside loadData */ });
  }, [loadData]);

  const handleToggle = React.useCallback((id: number): void => {
    setExpandedIds(prev => {
      const next = new Set(prev);
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }
      return next;
    });
  }, []);

  const openAddPanel = (): void => {
    setEditingFaq(undefined);
    setSubmitError(undefined);
    setIsPanelOpen(true);
  };

  const openEditPanel = (faq: IFAQ): void => {
    setEditingFaq(faq);
    setSubmitError(undefined);
    setIsPanelOpen(true);
  };

  const closePanel = (): void => {
    setIsPanelOpen(false);
    setEditingFaq(undefined);
    setSubmitError(undefined);
  };

  const handleSubmit = async (formData: IFAQFormData): Promise<void> => {
    setSubmitting(true);
    setSubmitError(undefined);
    try {
      if (editingFaq) {
        await service.updateFAQ(editingFaq.id, formData);
      } else {
        await service.addFAQ(formData);
      }
      closePanel();
      await loadData();
    } catch (err) {
      setSubmitError(
        err instanceof Error ? err.message : 'Failed to save FAQ. Please try again.'
      );
      // Error stays displayed in the panel — do not close
    } finally {
      setSubmitting(false);
    }
  };

  const handleDelete = async (id: number): Promise<void> => {
    if (!window.confirm('Are you sure you want to delete this FAQ?')) return;
      setError(undefined);
    try {
      await service.deleteFAQ(id);
      setExpandedIds(prev => {
        const next = new Set(prev);
        next.delete(id);
        return next;
      });
      await loadData();
    } catch (err) {
      setError(
        err instanceof Error ? err.message : 'Failed to delete FAQ. Please try again.'
      );
    }
  };

  if (loading) {
    return (
      <div className={styles.faqSection}>
        <Stack horizontalAlign="center" verticalAlign="center" style={{ minHeight: 160 }}>
          <Spinner size={SpinnerSize.large} label="Loading FAQs..." />
        </Stack>
      </div>
    );
  }

  return (
    <div className={styles.faqSection}>
      <div className={styles.header}>
        <Text as="h2" block>Frequently Asked Questions</Text>
        {canEdit && (
          <PrimaryButton
            iconProps={{ iconName: 'Add' }}
            text="Add FAQ"
            onClick={openAddPanel}
          />
        )}
      </div>

      {error && (
        <MessageBar
          messageBarType={MessageBarType.error}
          isMultiline={false}
          onDismiss={() => setError(undefined)}
          dismissButtonAriaLabel="Close"
          styles={{ root: { marginBottom: 16 } }}
        >
          {error}
        </MessageBar>
      )}

      {faqs.length === 0 && !error ? (
        <div className={styles.emptyState}>
          <Text>No FAQs available at this time.</Text>
        </div>
      ) : (
        <FAQList
          faqs={faqs}
          canEdit={canEdit}
          expandedIds={expandedIds}
          onToggle={handleToggle}
          onEdit={openEditPanel}
          onDelete={handleDelete}
        />
      )}

      <FAQForm
        isOpen={isPanelOpen}
        initialValues={editingFaq}
        submitting={submitting}
        submitError={submitError}
        onSubmit={handleSubmit}
        onDismiss={closePanel}
      />
    </div>
  );
};

export default FAQSection;
