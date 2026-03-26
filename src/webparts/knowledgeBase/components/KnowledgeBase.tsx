import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

import importedStyles from './KnowledgeBase.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;

import { IKnowledgeBaseProps } from './IKnowledgeBaseProps';
import { IKBArticle, KnowledgeBaseService } from '../../../services/KnowledgeBaseService';

import SearchBar from './SearchBar';
import CategoryFilter from './CategoryFilter';
import ArticleList from './ArticleList';
import ArticlePanel from './ArticlePanel';

const KnowledgeBase: React.FC<IKnowledgeBaseProps> = (props) => {
  const {
    title,
    listName,
    defaultCategory,
    itemsPerPage,
    enableSearch,
    enableCategoryFilter,
    spHttpClient,
    siteUrl
  } = props;

  // ── State ─────────────────────────────────────────────────────────
  const [articles, setArticles] = useState<IKBArticle[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | undefined>(undefined);

  const [searchQuery, setSearchQuery] = useState<string>('');
  const [selectedCategory, setSelectedCategory] = useState<string>(defaultCategory || 'All');

  const [selectedArticle, setSelectedArticle] = useState<IKBArticle | undefined>(undefined);
  const [isPanelOpen, setIsPanelOpen] = useState<boolean>(false);

  // ── Data fetching ─────────────────────────────────────────────────
  useEffect(() => {
    if (!listName) {
      setError('Please configure a SharePoint list name in the web part properties.');
      setLoading(false);
      return;
    }

    let cancelled = false;

    const fetchArticles = async (): Promise<void> => {
      setLoading(true);
      setError(undefined);
      try {
        const data = await KnowledgeBaseService.getArticles(siteUrl, listName, spHttpClient);
        if (!cancelled) {
          setArticles(data);
        }
      } catch (err) {
        if (!cancelled) {
          setError(
            err instanceof Error
              ? err.message
              : 'An unexpected error occurred while loading articles.'
          );
        }
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    };

    fetchArticles().catch(console.error);

    return () => {
      cancelled = true;
    };
  }, [siteUrl, listName]);

  // Reset category to default when it changes in property pane
  useEffect(() => {
    setSelectedCategory(defaultCategory || 'All');
  }, [defaultCategory]);

  // ── Derived data ──────────────────────────────────────────────────
  const categories = useMemo(
    () => KnowledgeBaseService.getDistinctCategories(articles),
    [articles]
  );

  const filteredArticles = useMemo(() => {
    let result = articles;

    // Category filter
    if (selectedCategory && selectedCategory !== 'All') {
      result = result.filter((a) => a.Category === selectedCategory);
    }

    // Search filter
    if (searchQuery.trim()) {
      const lowerQuery = searchQuery.toLowerCase();
      result = result.filter(
        (a) =>
          a.Title.toLowerCase().includes(lowerQuery) ||
          a.Description.toLowerCase().includes(lowerQuery)
      );
    }

    return result;
  }, [articles, selectedCategory, searchQuery]);

  // ── Handlers ──────────────────────────────────────────────────────
  const handleArticleClick = (article: IKBArticle): void => {
    setSelectedArticle(article);
    setIsPanelOpen(true);
  };

  const handlePanelDismiss = (): void => {
    setIsPanelOpen(false);
    setSelectedArticle(undefined);
  };

  // ── Render ────────────────────────────────────────────────────────
  if (loading) {
    return (
      <div className={`${styles.knowledgeBase} ${styles.loadingContainer}`}>
        <Spinner size={SpinnerSize.large} label="Loading articles..." />
      </div>
    );
  }

  return (
    <div className={styles.knowledgeBase}>
      {/* Header */}
      <div className={styles.header}>
        <Text as="h2" variant="xLarge" block styles={{ root: { margin: '0 0 4px 0', fontWeight: 600 } }}>
          {title}
        </Text>
        <Text className={styles.subtitle} block>
          {filteredArticles.length} article{filteredArticles.length !== 1 ? 's' : ''} found
        </Text>
      </div>

      {/* Error state */}
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

      {/* Controls: Search + Category Filter */}
      {!error && (
        <Stack className={styles.controls} tokens={{ childrenGap: 12 }}>
          {enableSearch && (
            <SearchBar value={searchQuery} onChange={setSearchQuery} />
          )}
          {enableCategoryFilter && categories.length > 0 && (
            <CategoryFilter
              categories={categories}
              selectedCategory={selectedCategory}
              onCategoryChange={setSelectedCategory}
            />
          )}
        </Stack>
      )}

      {/* Article List */}
      {!error && (
        <ArticleList
          articles={filteredArticles}
          searchQuery={searchQuery}
          itemsPerPage={itemsPerPage || 10}
          onArticleClick={handleArticleClick}
        />
      )}

      {/* Article Detail Panel */}
      <ArticlePanel
        article={selectedArticle}
        isOpen={isPanelOpen}
        onDismiss={handlePanelDismiss}
      />
    </div>
  );
};

export default KnowledgeBase;
