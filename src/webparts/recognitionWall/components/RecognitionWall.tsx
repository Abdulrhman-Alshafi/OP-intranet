import * as React from 'react';
import importedStyles from './RecognitionWall.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { IRecognitionWallProps } from './IRecognitionWallProps';
import { IKudosPost, IComment, IEmployeeOfMonth, RecognitionService } from '../../../services/RecognitionService';
import { escape } from '@microsoft/sp-lodash-subset';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { PrimaryButton, DefaultButton, IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import EmployeeHighlight from './EmployeeHighlight';
import KudosPost from './KudosPost';
import KudosForm from './KudosForm';
import EmployeeOfMonthForm from './EmployeeOfMonthForm';

/**
 * RecognitionWall is the main container component.
 * Manages data fetching, search, pagination, and polling.
 */
const RecognitionWall = (props: IRecognitionWallProps): React.ReactElement => {
  const { title, hasTeamsContext, showEmployeeOfMonth, siteUrl, isAdmin, postsPerPage, currentUserId, spHttpClient } = props;
  const serviceRef = React.useRef<RecognitionService>(new RecognitionService(spHttpClient, siteUrl));
  const currentPageRef = React.useRef<number>(0);
  const [posts, setPosts] = React.useState<IKudosPost[]>([]);
  const [employeeOfMonth, setEmployeeOfMonth] = React.useState<IEmployeeOfMonth | null>(null);
  const [likeStatuses, setLikeStatuses] = React.useState<{ [postId: number]: boolean }>({});
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | null>(null);
  const [searchQuery, setSearchQuery] = React.useState<string>('');
  const [showForm, setShowForm] = React.useState<boolean>(false);
  const [showEotmForm, setShowEotmForm] = React.useState<boolean>(false);
  const [loadingMore, setLoadingMore] = React.useState<boolean>(false);
  const [hasMore, setHasMore] = React.useState<boolean>(true);
  const [lastUpdated, setLastUpdated] = React.useState<Date | null>(null);

  React.useEffect(() => {
    serviceRef.current = new RecognitionService(props.spHttpClient, props.siteUrl);
  }, [props.spHttpClient, props.siteUrl]);

  const loadInitialData = async (): Promise<void> => {
    setLoading(true);
    setError(null);
    currentPageRef.current = 0;

    try {
      const { posts: loadedPosts } = await serviceRef.current.getPosts(postsPerPage, 0);
      const loadedEmployeeOfMonth = showEmployeeOfMonth ? await serviceRef.current.getEmployeeOfMonth() : null;
      const loadedLikeStatuses = currentUserId > 0
        ? await serviceRef.current.getUserLikeStatuses(loadedPosts.map((post) => post.Id), currentUserId)
        : {};

      setPosts(loadedPosts);
      setEmployeeOfMonth(loadedEmployeeOfMonth);
      setLikeStatuses(loadedLikeStatuses);
      setLoading(false);
      setHasMore(loadedPosts.length >= postsPerPage);
      setLastUpdated(new Date());
    } catch (err) {
      console.error('RecognitionWall: Failed to load data', err);
      setLoading(false);
      setError((err as Error).message || 'Failed to load recognition data.');
    }
  };

  React.useEffect(() => {
    loadInitialData().catch((err) => {
      console.error('RecognitionWall: Failed to load data', err);
    });
  }, [postsPerPage, showEmployeeOfMonth, currentUserId, props.siteUrl, props.spHttpClient]);

  const loadMore = async (): Promise<void> => {
    setLoadingMore(true);
    currentPageRef.current += 1;

    try {
      const skip = currentPageRef.current * postsPerPage;
      const { posts: newPosts } = await serviceRef.current.getPosts(postsPerPage, skip, searchQuery || undefined);

      if (currentUserId > 0 && newPosts.length > 0) {
        const newLikes = await serviceRef.current.getUserLikeStatuses(newPosts.map((post) => post.Id), currentUserId);
        setPosts((currentPosts) => [...currentPosts, ...newPosts]);
        setLikeStatuses((currentLikeStatuses) => ({ ...currentLikeStatuses, ...newLikes }));
        setLoadingMore(false);
        setHasMore(newPosts.length >= postsPerPage);
      } else {
        setPosts((currentPosts) => [...currentPosts, ...newPosts]);
        setLoadingMore(false);
        setHasMore(newPosts.length >= postsPerPage);
      }
    } catch {
      setLoadingMore(false);
    }
  };

  const handleSearch = async (query: string): Promise<void> => {
    setSearchQuery(query);
    setLoading(true);
    currentPageRef.current = 0;

    try {
      const { posts: searchedPosts } = await serviceRef.current.getPosts(postsPerPage, 0, query || undefined);
      const searchedLikeStatuses = currentUserId > 0
        ? await serviceRef.current.getUserLikeStatuses(searchedPosts.map((post) => post.Id), currentUserId)
        : {};

      setPosts(searchedPosts);
      setLikeStatuses(searchedLikeStatuses);
      setLoading(false);
      setHasMore(searchedPosts.length >= postsPerPage);
    } catch {
      setLoading(false);
    }
  };

  const handleToggleLike = async (postId: number): Promise<void> => {
    if (currentUserId === 0) {
      return;
    }

    try {
      const isNowLiked = await serviceRef.current.toggleLike(postId, currentUserId);

      setPosts((currentPosts) => currentPosts.map((post) => {
        if (post.Id === postId) {
          return { ...post, LikesCount: post.LikesCount + (isNowLiked ? 1 : -1) };
        }
        return post;
      }));
      setLikeStatuses((currentLikeStatuses) => ({ ...currentLikeStatuses, [postId]: isNowLiked }));
    } catch (err) {
      console.error('RecognitionWall: Failed to toggle like', err);
    }
  };

  const handleLoadComments = async (postId: number): Promise<IComment[]> => {
    return serviceRef.current.getComments(postId);
  };

  const handleAddComment = async (postId: number, text: string): Promise<void> => {
    setPosts((currentPosts) => currentPosts.map((post) =>
      post.Id === postId ? { ...post, CommentsCount: (post.CommentsCount || 0) + 1 } : post
    ));

    try {
      await serviceRef.current.addComment(postId, text);
    } catch (err) {
      console.error('RecognitionWall: Failed to add comment', err);
      setPosts((currentPosts) => currentPosts.map((post) =>
        post.Id === postId ? { ...post, CommentsCount: Math.max(0, (post.CommentsCount || 0) - 1) } : post
      ));
    }
  };

  const handleKudosSubmit = async (recipientEmail: string, _recipientName: string, message: string, category: string): Promise<void> => {
    try {
      await serviceRef.current.createPost(recipientEmail, message, category);
      setShowForm(false);
      await loadInitialData();
    } catch (err) {
      console.error('RecognitionWall: Failed to create kudos', err);
    }
  };

  const handleDeletePost = async (postId: number): Promise<void> => {
    try {
      await serviceRef.current.deletePost(postId);
      setPosts((currentPosts) => currentPosts.filter((post) => post.Id !== postId));
    } catch (err) {
      console.error('RecognitionWall: Failed to delete post', err);
    }
  };

  const handleSetEmployee = async (email: string, name: string): Promise<void> => {
    try {
      await serviceRef.current.setEmployeeOfMonth(email, name);
      setShowEotmForm(false);
      await loadInitialData();
    } catch (err) {
      console.error('RecognitionWall: Failed to set Employee of the Month', err);
    }
  };

  const handleRemoveEmployee = async (): Promise<void> => {
    try {
      await serviceRef.current.removeEmployeeOfMonth();
      setEmployeeOfMonth(null);
    } catch (err) {
      console.error('RecognitionWall: Failed to remove Employee of the Month', err);
    }
  };

  const handleRefresh = async (): Promise<void> => {
    await loadInitialData();
  };

  return (
    <section className={`${styles.recognitionWall} ${hasTeamsContext ? styles.teams : ''}`}>
        {/* Header */}
        <div className={styles.header}>
          <div className={styles.titleSection}>
            <h2>{escape(title)}</h2>
            {!loading && (
              <div className={styles.summary}>
                <div className={styles.summaryItem}>
                  <Icon iconName="FavoriteStar" className={styles.icon} />
                  <span>{posts.length} kudos</span>
                </div>
              </div>
            )}
          </div>
          <div className={styles.controls}>
            <SearchBox
              placeholder="Search kudos..."
              onSearch={(val) => {
                handleSearch(val || '').catch(() => {
                  setLoading(false);
                });
              }}
              onClear={() => {
                handleSearch('').catch(() => {
                  setLoading(false);
                });
              }}
              className={styles.searchBox}
              styles={{ root: { minWidth: 200 } }}
            />
            {isAdmin && (
              <PrimaryButton
                text="Give Kudos"
                iconProps={{ iconName: 'FavoriteStar' }}
                onClick={() => setShowForm(true)}
              />
            )}
            {isAdmin && showEmployeeOfMonth && (
              <DefaultButton
                text="Employee of the Month"
                iconProps={{ iconName: 'Ribbon' }}
                className={styles.eotmButton}
                onClick={() => setShowEotmForm(true)}
              />
            )}
            <IconButton
              iconProps={{ iconName: 'Refresh' }}
              title="Refresh"
              ariaLabel="Refresh"
              onClick={() => {
                handleRefresh().catch((err) => {
                  console.error('RecognitionWall: Refresh failed', err);
                });
              }}
              disabled={loading}
            />
            {lastUpdated && (
              <div className={styles.lastUpdated}>
                Updated: {lastUpdated.toLocaleTimeString()}
              </div>
            )}
          </div>
        </div>

        {/* Error */}
        {error && (
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={false}
            onDismiss={() => setError(null)}
          >
            {error}
          </MessageBar>
        )}

        {/* Employee of the Month */}
        {showEmployeeOfMonth && employeeOfMonth && !loading && (
          <EmployeeHighlight
            employee={employeeOfMonth}
            siteUrl={siteUrl}
            isAdmin={isAdmin}
            onRemove={() => {
              handleRemoveEmployee().catch((err) => {
                console.error('RecognitionWall: Failed to remove Employee of the Month', err);
              });
            }}
          />
        )}

        {/* Loading state */}
        {loading && (
          <div className={styles.loadingContainer}>
            <Spinner size={SpinnerSize.large} label="Loading recognition wall..." />
          </div>
        )}

        {/* Empty state */}
        {!loading && posts.length === 0 && !error && (
          <div className={styles.emptyState}>
            <Icon iconName="FavoriteStar" className={styles.emptyIcon} />
            <h3>No kudos yet</h3>
            <p>Be the first to recognise a colleague! Click &quot;Give Kudos&quot; to get started.</p>
          </div>
        )}

        {/* Posts feed */}
        {!loading && posts.length > 0 && (
          <div className={styles.feed}>
            {posts.map(post => (
              <KudosPost
                key={post.Id}
                post={post}
                currentUserId={currentUserId}
                hasLiked={likeStatuses[post.Id] || false}
                isAdmin={isAdmin}
                onToggleLike={handleToggleLike}
                onAddComment={handleAddComment}
                onLoadComments={handleLoadComments}
                onDelete={handleDeletePost}
              />
            ))}
          </div>
        )}

        {/* Load more */}
        {!loading && hasMore && posts.length > 0 && (
          <div className={styles.loadMoreContainer}>
            {loadingMore ? (
              <Spinner size={SpinnerSize.small} label="Loading more..." />
            ) : (
              <button
                className={styles.loadMoreBtn}
                onClick={() => {
                  loadMore().catch(() => {
                    setLoadingMore(false);
                  });
                }}
              >
                Load More
              </button>
            )}
          </div>
        )}

        {/* Kudos Form Panel */}
        <KudosForm
          isOpen={showForm}
          onDismiss={() => setShowForm(false)}
          onSubmit={handleKudosSubmit}
          siteUrl={props.siteUrl}
          spHttpClient={props.spHttpClient}
        />

        {/* Employee of the Month Form Panel */}
        <EmployeeOfMonthForm
          isOpen={showEotmForm}
          onDismiss={() => setShowEotmForm(false)}
          onSubmit={handleSetEmployee}
          siteUrl={props.siteUrl}
          spHttpClient={props.spHttpClient}
        />
      </section>
  );
};

export default RecognitionWall;
