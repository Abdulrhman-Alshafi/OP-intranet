import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { ActionButton } from '@fluentui/react/lib/Button';
import { Text } from '@fluentui/react/lib/Text';

import { IFAQ } from '../../../services/FAQService';
import importedStyles from './FAQSection.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;

interface IFAQItemProps {
  faq: IFAQ;
  isExpanded: boolean;
  canEdit: boolean;
  onToggle: (id: number) => void;
  onEdit: (faq: IFAQ) => void;
  onDelete: (id: number) => void;
}

export const FAQItem = ({
  faq,
  isExpanded,
  canEdit,
  onToggle,
  onEdit,
  onDelete
}: IFAQItemProps): JSX.Element => {
  const contentRef = React.useRef<HTMLDivElement>(null);

  // Derive max-height for smooth animation: measure the inner div height when expanded
  const maxHeight = React.useMemo<string>(() => {
    if (!isExpanded) return '0px';
    if (contentRef.current) {
      return `${contentRef.current.scrollHeight + 8}px`;
    }
    // Fallback before first render measurement
    return '2000px';
  }, [isExpanded]);

  const handleHeaderClick = (): void => {
    onToggle(faq.id);
  };

  const handleHeaderKeyDown = (e: React.KeyboardEvent<HTMLDivElement>): void => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      onToggle(faq.id);
    }
  };

  const handleEditClick = (): void => {
    onEdit(faq);
  };

  const handleDeleteClick = (): void => {
    onDelete(faq.id);
  };

  return (
    <div className={styles.faqItem}>
      {/* ── Clickable header ─────────────────────────────── */}
      <div
        className={styles.faqHeader}
        onClick={handleHeaderClick}
        onKeyDown={handleHeaderKeyDown}
        role="button"
        tabIndex={0}
        aria-expanded={isExpanded}
        aria-controls={`faq-answer-${faq.id}`}
      >
        <div className={styles.faqHeaderLeft}>
          <Icon
            iconName="ChevronRight"
            className={`${styles.chevron} ${isExpanded ? styles.chevronExpanded : ''}`}
            aria-hidden
          />
          <Text className={styles.questionText}>{faq.title}</Text>
        </div>

        {canEdit && (
          // stopPropagation on the wrapper prevents the header toggle from firing
          // eslint-disable-next-line react/jsx-no-bind
          <div className={styles.adminActions} onClick={e => e.stopPropagation()}>
            <ActionButton
              iconProps={{ iconName: 'Edit' }}
              title="Edit FAQ"
              ariaLabel="Edit FAQ"
              onClick={handleEditClick}
              styles={{ root: { height: 28, minWidth: 0, padding: '0 6px' } }}
            />
            <ActionButton
              iconProps={{ iconName: 'Delete' }}
              title="Delete FAQ"
              ariaLabel="Delete FAQ"
              onClick={handleDeleteClick}
              styles={{
                root: { height: 28, minWidth: 0, padding: '0 6px' },
                icon: { color: '#a4262c' }
              }}
            />
          </div>
        )}
      </div>

      {/* ── Animated answer body ──────────────────────────── */}
      <div
        id={`faq-answer-${faq.id}`}
        ref={contentRef}
        role="region"
        aria-labelledby={`faq-question-${faq.id}`}
        className={`${styles.answerBody} ${isExpanded ? styles.answerBodyExpanded : ''}`}
        style={{ maxHeight }}
      >
        <div className={styles.answerContent}>{faq.answer}</div>
      </div>
    </div>
  );
};

export default FAQItem;
