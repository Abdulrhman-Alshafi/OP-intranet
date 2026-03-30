import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';

import { IFAQ } from '../../../services/FAQService';
import { FAQItem } from './FAQItem';
import importedStyles from './FAQSection.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;

interface IFAQListProps {
  faqs: IFAQ[];
  canEdit: boolean;
  expandedIds: Set<number>;
  onToggle: (id: number) => void;
  onEdit: (faq: IFAQ) => void;
  onDelete: (id: number) => void;
}

interface IFAQGroup {
  category: string;
  items: IFAQ[];
}

export const FAQList = ({
  faqs,
  canEdit,
  expandedIds,
  onToggle,
  onEdit,
  onDelete
}: IFAQListProps): JSX.Element => {
  const groups = React.useMemo<IFAQGroup[]>(() => {
    const map: Map<string, IFAQ[]> = new Map();

    faqs.forEach(faq => {
      const key = faq.category && faq.category.trim() ? faq.category.trim() : 'General';
      if (!map.has(key)) {
        map.set(key, []);
      }
      // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
      map.get(key)!.push(faq);
    });

    return Array.from(map.entries()).map(([category, items]) => ({ category, items }));
  }, [faqs]);

  return (
    <Stack tokens={{ childrenGap: 0 }}>
      {groups.map(group => (
        <div key={group.category} className={styles.categoryGroup}>
          <div className={styles.categoryLabel}>
            <Text>{group.category}</Text>
          </div>
          {group.items.map(faq => (
            <FAQItem
              key={faq.id}
              faq={faq}
              isExpanded={expandedIds.has(faq.id)}
              canEdit={canEdit}
              onToggle={onToggle}
              onEdit={onEdit}
              onDelete={onDelete}
            />
          ))}
        </div>
      ))}
    </Stack>
  );
};

export default FAQList;
