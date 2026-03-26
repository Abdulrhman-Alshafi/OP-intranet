import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Text } from '@fluentui/react/lib/Text';
import { Icon } from '@fluentui/react/lib/Icon';
import importedStyles from './KnowledgeBase.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { IKBArticle } from '../../../services/KnowledgeBaseService';
import ArticleCard from './ArticleCard';

interface IArticleListProps {
  articles: IKBArticle[];
  searchQuery: string;
  itemsPerPage: number;
  onArticleClick: (article: IKBArticle) => void;
}

const ArticleList: React.FC<IArticleListProps> = ({
  articles,
  searchQuery,
  itemsPerPage,
  onArticleClick
}) => {
  const [currentPage, setCurrentPage] = React.useState<number>(1);

  // Reset to page 1 when articles list changes (search/filter)
  React.useEffect(() => {
    setCurrentPage(1);
  }, [articles]);

  const totalPages = Math.max(1, Math.ceil(articles.length / itemsPerPage));
  const startIndex = (currentPage - 1) * itemsPerPage;
  const visibleArticles = articles.slice(startIndex, startIndex + itemsPerPage);

  if (articles.length === 0) {
    return (
      <div className={styles.emptyState}>
        <div className={styles.emptyIcon}>
          <Icon iconName="KnowledgeArticle" />
        </div>
        <p>No articles found. Try adjusting your search or category filter.</p>
      </div>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 0 }}>
      <div className={styles.articleGrid}>
        {visibleArticles.map((article) => (
          <ArticleCard
            key={article.Id}
            article={article}
            searchQuery={searchQuery}
            onClick={onArticleClick}
          />
        ))}
      </div>

      {totalPages > 1 && (
        <div className={styles.pagination}>
          <DefaultButton
            text="Previous"
            iconProps={{ iconName: 'ChevronLeft' }}
            disabled={currentPage === 1}
            onClick={() => setCurrentPage((p) => p - 1)}
          />
          <Text className={styles.pageInfo}>
            Page {currentPage} of {totalPages}
          </Text>
          <DefaultButton
            text="Next"
            iconProps={{ iconName: 'ChevronRight' }}
            disabled={currentPage === totalPages}
            onClick={() => setCurrentPage((p) => p + 1)}
          />
        </div>
      )}
    </Stack>
  );
};

export default ArticleList;
