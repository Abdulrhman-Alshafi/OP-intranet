import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import importedStyles from './KnowledgeBase.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { IKBArticle } from '../../../services/KnowledgeBaseService';

interface IArticleCardProps {
  article: IKBArticle;
  searchQuery: string;
  onClick: (article: IKBArticle) => void;
}

/**
 * Wraps matched text in a <mark> span for visual highlighting.
 */
const HighlightText: React.FC<{ text: string; query: string }> = ({ text, query }) => {
  if (!query.trim()) {
    return <>{text}</>;
  }

  const escapedQuery = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  // eslint-disable-next-line @rushstack/security/no-unsafe-regexp
  const splitRegex = new RegExp(`(${escapedQuery})`, 'gi');
  const parts = text.split(splitRegex);

  return (
    <>
      {parts.map((part, index) =>
        part.toLowerCase() === query.toLowerCase() ? (
          <span key={index} className={styles.highlight}>{part}</span>
        ) : (
          <React.Fragment key={index}>{part}</React.Fragment>
        )
      )}
    </>
  );
};

const ArticleCard: React.FC<IArticleCardProps> = ({ article, searchQuery, onClick }) => {
  const formattedDate = React.useMemo(() => {
    if (!article.Created) return '';
    try {
      return new Date(article.Created).toLocaleDateString(undefined, {
        year: 'numeric',
        month: 'short',
        day: 'numeric'
      });
    } catch {
      return article.Created;
    }
  }, [article.Created]);

  return (
    <div
      className={styles.articleCard}
      onClick={() => onClick(article)}
      role="button"
      tabIndex={0}
      onKeyDown={(e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault();
          onClick(article);
        }
      }}
      aria-label={`Open article: ${article.Title}`}
    >
      <Stack tokens={{ childrenGap: 0 }}>
        <Text className={styles.cardTitle} block>
          <HighlightText text={article.Title} query={searchQuery} />
        </Text>

        {article.Description && (
          <Text className={styles.cardDescription} block>
            <HighlightText text={article.Description} query={searchQuery} />
          </Text>
        )}

        <div className={styles.cardMeta}>
          <span className={styles.categoryBadge}>{article.Category}</span>
          <span className={styles.cardDate}>{formattedDate}</span>
        </div>
      </Stack>
    </div>
  );
};

export default ArticleCard;
