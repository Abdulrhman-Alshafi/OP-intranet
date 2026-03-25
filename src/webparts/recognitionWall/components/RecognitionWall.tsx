import * as React from 'react';
import importedStyles from './RecognitionWall.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { IRecognitionWallProps } from './IRecognitionWallProps';
import { IKudosPost, ITopEmployee, IComment, IEmployeeOfMonth, RecognitionService } from '../../../services/RecognitionService';
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

export interface IRecognitionWallState {
  posts: IKudosPost[];
  employeeOfMonth: IEmployeeOfMonth | null;
  likeStatuses: { [postId: number]: boolean };
  loading: boolean;
  error: string | null;
  searchQuery: string;
  showForm: boolean;
  showEotmForm: boolean;
  loadingMore: boolean;
  hasMore: boolean;
  lastUpdated: Date | null;
}

/**
 * RecognitionWall is the main container component.
 * Manages data fetching, search, pagination, and polling.
 */
export default class RecognitionWall extends React.Component<IRecognitionWallProps, IRecognitionWallState> {
  private _service: RecognitionService;
  private _currentPage: number = 0;

  constructor(props: IRecognitionWallProps) {
    super(props);

    this._service = new RecognitionService(props.spHttpClient, props.siteUrl);

    this.state = {
      posts: [],
      employeeOfMonth: null,
      likeStatuses: {},
      loading: true,
      error: null,
      searchQuery: '',
      showForm: false,
      showEotmForm: false,
      loadingMore: false,
      hasMore: true,
      lastUpdated: null
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._loadInitialData();
  }

  /**
   * Load posts, employee of the month, and like statuses.
   */
  private async _loadInitialData(): Promise<void> {
    this.setState({ loading: true, error: null });
    this._currentPage = 0;

    try {
      const { posts } = await this._service.getPosts(this.props.postsPerPage, 0);
      const employeeOfMonth = this.props.showEmployeeOfMonth ? await this._service.getEmployeeOfMonth() : null;
      const likeStatuses = this.props.currentUserId > 0
        ? await this._service.getUserLikeStatuses(posts.map(p => p.Id), this.props.currentUserId)
        : {};

      this.setState({
        posts,
        employeeOfMonth,
        likeStatuses,
        loading: false,
        hasMore: posts.length >= this.props.postsPerPage,
        lastUpdated: new Date()
      });
    } catch (err) {
      console.error('RecognitionWall: Failed to load data', err);
      this.setState({
        loading: false,
        error: (err as Error).message || 'Failed to load recognition data.'
      });
    }
  }

  /**
   * Load more posts (pagination via "Load More" button).
   */
  private _loadMore = async (): Promise<void> => {
    this.setState({ loadingMore: true });
    this._currentPage++;

    try {
      const skip = this._currentPage * this.props.postsPerPage;
      const { posts: newPosts } = await this._service.getPosts(this.props.postsPerPage, skip, this.state.searchQuery || undefined);

      if (this.props.currentUserId > 0 && newPosts.length > 0) {
        const newLikes = await this._service.getUserLikeStatuses(newPosts.map(p => p.Id), this.props.currentUserId);
        this.setState(prev => ({
          posts: [...prev.posts, ...newPosts],
          likeStatuses: { ...prev.likeStatuses, ...newLikes },
          loadingMore: false,
          hasMore: newPosts.length >= this.props.postsPerPage
        }));
      } else {
        this.setState(prev => ({
          posts: [...prev.posts, ...newPosts],
          loadingMore: false,
          hasMore: newPosts.length >= this.props.postsPerPage
        }));
      }
    } catch {
      this.setState({ loadingMore: false });
    }
  };

  /**
   * Handle search query change.
   */
  private _handleSearch = async (query: string): Promise<void> => {
    this.setState({ searchQuery: query, loading: true });
    this._currentPage = 0;

    try {
      const { posts } = await this._service.getPosts(this.props.postsPerPage, 0, query || undefined);
      const likeStatuses = this.props.currentUserId > 0
        ? await this._service.getUserLikeStatuses(posts.map(p => p.Id), this.props.currentUserId)
        : {};

      this.setState({
        posts,
        likeStatuses,
        loading: false,
        hasMore: posts.length >= this.props.postsPerPage
      });
    } catch {
      this.setState({ loading: false });
    }
  };

  /**
   * Toggle like on a post.
   */
  private _handleToggleLike = async (postId: number): Promise<void> => {
    if (this.props.currentUserId === 0) return;

    try {
      const isNowLiked = await this._service.toggleLike(postId, this.props.currentUserId);

      this.setState(prev => {
        const updatedPosts = prev.posts.map(p => {
          if (p.Id === postId) {
            return { ...p, LikesCount: p.LikesCount + (isNowLiked ? 1 : -1) };
          }
          return p;
        });
        return {
          posts: updatedPosts,
          likeStatuses: { ...prev.likeStatuses, [postId]: isNowLiked }
        };
      });
    } catch (err) {
      console.error('RecognitionWall: Failed to toggle like', err);
    }
  };

  /**
   * Load comments for a post.
   */
  private _handleLoadComments = async (postId: number): Promise<IComment[]> => {
    return this._service.getComments(postId);
  };

  /**
   * Add a comment to a post.
   */
  private _handleAddComment = async (postId: number, text: string): Promise<void> => {
    // Optimistically update comment count immediately
    this.setState(prev => ({
      posts: prev.posts.map(p =>
        p.Id === postId ? { ...p, CommentsCount: (p.CommentsCount || 0) + 1 } : p
      )
    }));

    try {
      await this._service.addComment(postId, text);
    } catch (err) {
      console.error('RecognitionWall: Failed to add comment', err);
      // Rollback the optimistic update on error
      this.setState(prev => ({
        posts: prev.posts.map(p =>
          p.Id === postId ? { ...p, CommentsCount: Math.max(0, (p.CommentsCount || 0) - 1) } : p
        )
      }));
    }
  };

  /**
   * Handle new kudos submission from the form.
   */
  private _handleKudosSubmit = async (recipientEmail: string, _recipientName: string, message: string, category: string): Promise<void> => {
    try {
      await this._service.createPost(recipientEmail, message, category);
      this.setState({ showForm: false });
      // Reload to show the new post
      await this._loadInitialData();
    } catch (err) {
      console.error('RecognitionWall: Failed to create kudos', err);
    }
  };

  /**
   * Handle deleting a kudos post (admin only).
   */
  private _handleDeletePost = async (postId: number): Promise<void> => {
    try {
      await this._service.deletePost(postId);
      // Remove the post from state immediately
      this.setState(prev => ({
        posts: prev.posts.filter(p => p.Id !== postId)
      }));
    } catch (err) {
      console.error('RecognitionWall: Failed to delete post', err);
    }
  };

  /**
   * Handle setting Employee of the Month (admin only).
   */
  private _handleSetEmployee = async (email: string, name: string): Promise<void> => {
    try {
      await this._service.setEmployeeOfMonth(email, name);
      this.setState({ showEotmForm: false });
      await this._loadInitialData();
    } catch (err) {
      console.error('RecognitionWall: Failed to set Employee of the Month', err);
    }
  };

  /**
   * Handle removing Employee of the Month (admin only).
   */
  private _handleRemoveEmployee = async (): Promise<void> => {
    try {
      await this._service.removeEmployeeOfMonth();
      this.setState({ employeeOfMonth: null });
    } catch (err) {
      console.error('RecognitionWall: Failed to remove Employee of the Month', err);
    }
  };

  /**
   * Handle manual refresh.
   */
  private _handleRefresh = async (): Promise<void> => {
    await this._loadInitialData();
  };

  public render(): React.ReactElement<IRecognitionWallProps> {
    const { title, hasTeamsContext, showEmployeeOfMonth, siteUrl, isAdmin } = this.props;
    const { posts, employeeOfMonth, likeStatuses, loading, error, showForm, showEotmForm, loadingMore, hasMore, lastUpdated } = this.state;

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
              onSearch={(val) => { this._handleSearch(val || '').catch(console.error); }}
              onClear={() => { this._handleSearch('').catch(console.error); }}
              className={styles.searchBox}
              styles={{ root: { minWidth: 200 } }}
            />
            {isAdmin && (
              <PrimaryButton
                text="Give Kudos"
                iconProps={{ iconName: 'FavoriteStar' }}
                onClick={() => this.setState({ showForm: true })}
              />
            )}
            {isAdmin && showEmployeeOfMonth && (
              <DefaultButton
                text="Employee of the Month"
                iconProps={{ iconName: 'Ribbon' }}
                className={styles.eotmButton}
                onClick={() => this.setState({ showEotmForm: true })}
              />
            )}
            <IconButton
              iconProps={{ iconName: 'Refresh' }}
              title="Refresh"
              ariaLabel="Refresh"
              onClick={() => { this._handleRefresh().catch(console.error); }}
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
            onDismiss={() => this.setState({ error: null })}
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
            onRemove={() => { this._handleRemoveEmployee().catch(console.error); }}
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
                currentUserId={this.props.currentUserId}
                hasLiked={likeStatuses[post.Id] || false}
                isAdmin={isAdmin}
                onToggleLike={this._handleToggleLike}
                onAddComment={this._handleAddComment}
                onLoadComments={this._handleLoadComments}
                onDelete={this._handleDeletePost}
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
              <button className={styles.loadMoreBtn} onClick={() => { this._loadMore().catch(console.error); }}>
                Load More
              </button>
            )}
          </div>
        )}

        {/* Kudos Form Panel */}
        <KudosForm
          isOpen={showForm}
          onDismiss={() => this.setState({ showForm: false })}
          onSubmit={this._handleKudosSubmit}
          siteUrl={this.props.siteUrl}
          spHttpClient={this.props.spHttpClient}
        />

        {/* Employee of the Month Form Panel */}
        <EmployeeOfMonthForm
          isOpen={showEotmForm}
          onDismiss={() => this.setState({ showEotmForm: false })}
          onSubmit={this._handleSetEmployee}
          siteUrl={this.props.siteUrl}
          spHttpClient={this.props.spHttpClient}
        />
      </section>
    );
  }
}
