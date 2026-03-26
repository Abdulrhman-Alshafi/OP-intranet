import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import importedStyles from './KnowledgeBase.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { IKBArticle } from '../../../services/KnowledgeBaseService';

interface IArticlePanelProps {
  article: IKBArticle | undefined;
  isOpen: boolean;
  onDismiss: () => void;
}

const ArticlePanel = ({ article, isOpen, onDismiss }: IArticlePanelProps): JSX.Element => {
  const formattedDate = React.useMemo(() => {
    if (!article?.Created) return '';
    try {
      return new Date(article.Created).toLocaleDateString(undefined, {
        weekday: 'long',
        year: 'numeric',
        month: 'long',
        day: 'numeric'
      });
    } catch {
      return article.Created;
    }
  }, [article?.Created]);

  if (!article) return <></>;

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText={article.Title}
      isLightDismiss
      styles={{
        header: { paddingTop: 24 },
        content: { paddingTop: 0 }
      }}
    >
      <div className={styles.panelContent}>
        <Stack className={styles.panelMeta} horizontal tokens={{ childrenGap: 12 }} wrap>
          <span className={styles.categoryBadge}>{article.Category}</span>
          <Text className={styles.panelDate}>{formattedDate}</Text>
        </Stack>

        {article.Content ? (
          <div
            className={styles.panelBody}
            // Content is from SharePoint rich text — render as HTML
            // eslint-disable-next-line react/no-danger
            dangerouslySetInnerHTML={{ __html: article.Content }}
          />
        ) : article.Description ? (
          <div className={styles.panelBody}>
            <p>{article.Description}</p>
          </div>
        ) : (
          <Text className={styles.panelDate}>No content available for this article.</Text>
        )}
      </div>
    </Panel>
  );
};

export default ArticlePanel;
