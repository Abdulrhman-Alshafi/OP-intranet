import * as React from 'react';
import importedStyles from './RecognitionWall.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { IKudosPostProps } from './IRecognitionWallProps';
import { IComment } from '../../../services/RecognitionService';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import CommentSection from './CommentSection';

export interface IKudosPostState {
  showComments: boolean;
  comments: IComment[];
  commentsLoading: boolean;
  liking: boolean;
  deleting: boolean;
}

/**
 * KudosPost renders a single kudos card with sender→recipient, message,
 * category badge, like button, and expandable comments.
 */
export default class KudosPost extends React.Component<IKudosPostProps, IKudosPostState> {
  constructor(props: IKudosPostProps) {
    super(props);
    this.state = {
      showComments: false,
      comments: [],
      commentsLoading: false,
      liking: false,
      deleting: false
    };
  }

  /**
   * Relative time formatting (e.g. "2h ago", "3d ago")
   */
  private _formatDate(dateStr: string): string {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    const now = new Date();
    const diffMs = now.getTime() - d.getTime();
    const diffMin = Math.floor(diffMs / 60000);
    const diffHr = Math.floor(diffMs / 3600000);
    const diffDays = Math.floor(diffMs / 86400000);

    if (diffMin < 1) return 'Just now';
    if (diffMin < 60) return `${diffMin}m ago`;
    if (diffHr < 24) return `${diffHr}h ago`;
    if (diffDays < 7) return `${diffDays}d ago`;
    return d.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' });
  }

  /**
   * Toggle the comments section, loading comments on first expand.
   */
  private _toggleComments = async (): Promise<void> => {
    const show = !this.state.showComments;
    this.setState({ showComments: show });

    if (show && this.state.comments.length === 0) {
      this.setState({ commentsLoading: true });
      try {
        const comments = await this.props.onLoadComments(this.props.post.Id);
        this.setState({ comments, commentsLoading: false });
      } catch {
        this.setState({ commentsLoading: false });
      }
    }
  };

  /**
   * Handle like button click
   */
  private _handleLike = async (): Promise<void> => {
    if (this.state.liking) return;
    this.setState({ liking: true });
    try {
      await this.props.onToggleLike(this.props.post.Id);
    } finally {
      // Reset liking state after the operation completes
      this.setState({ liking: false });
    }
  };

  /**
   * Handle a new comment being added (optimistic update)
   */
  private _handleAddComment = (postId: number, text: string): void => {
    this.props.onAddComment(postId, text);
    // Optimistic: add the comment locally
    const optimisticComment: IComment = {
      Id: Date.now(),
      PostId: postId,
      Comment: text,
      AuthorId: this.props.currentUserId,
      AuthorName: 'You',
      Created: new Date().toISOString()
    };
    this.setState(prev => ({
      comments: [...prev.comments, optimisticComment]
    }));
  };

  /**
   * Get the category color accent
   */
  private _getCategoryClass(): string {
    const cat = this.props.post.Category;
    switch (cat) {
      case 'Teamwork': return styles.catTeamwork;
      case 'Innovation': return styles.catInnovation;
      case 'Leadership': return styles.catLeadership;
      case 'Customer Focus': return styles.catCustomer;
      case 'Above & Beyond': return styles.catAboveBeyond;
      default: return styles.catDefault;
    }
  }

  /**
   * Handle delete button click (admin only)
   */
  private _handleDelete = async (): Promise<void> => {
    if (this.state.deleting) return;
    this.setState({ deleting: true });
    try {
      await this.props.onDelete(this.props.post.Id);
    } catch {
      this.setState({ deleting: false });
    }
  };

  public render(): React.ReactElement<IKudosPostProps> {
    const { post, hasLiked, isAdmin } = this.props;
    const { showComments, comments, commentsLoading, deleting } = this.state;

    return (
      <div className={styles.kudosCard}>
        {/* Admin delete button */}
        {isAdmin && (
          <IconButton
            iconProps={{ iconName: 'Cancel' }}
            title="Delete kudos"
            ariaLabel="Delete kudos"
            className={styles.deleteBtn}
            onClick={() => { this._handleDelete().catch(console.error); }}
            disabled={deleting}
          />
        )}
        {/* Card header: sender → recipient */}
        <div className={styles.kudosCardHeader}>
          <div className={styles.kudosSender}>
            <Icon iconName="Contact" className={styles.personIcon} />
            <span className={styles.personName}>{post.SenderName}</span>
          </div>
          <Icon iconName="Forward" className={styles.arrowIcon} />
          <div className={styles.kudosRecipient}>
            <Icon iconName="Contact" className={styles.personIcon} />
            <span className={styles.personName}>{post.RecipientName}</span>
          </div>
        </div>

        {/* Category badge */}
        {post.Category && (
          <span className={`${styles.categoryBadge} ${this._getCategoryClass()}`}>
            {post.Category}
          </span>
        )}

        {/* Message body */}
        <div className={styles.kudosMessage}>{post.Title}</div>

        {/* Date */}
        <div className={styles.kudosDate}>
          <Icon iconName="Clock" className={styles.clockIcon} />
          <span>{this._formatDate(post.Created)}</span>
        </div>

        {/* Actions: like + comments toggle */}
        <div className={styles.kudosActions}>
          <button
            className={`${styles.actionBtn} ${hasLiked ? styles.liked : ''}`}
            onClick={() => { this._handleLike().catch(console.error); }}
            disabled={this.state.liking}
            title={hasLiked ? 'Unlike' : 'Like'}
          >
            <Icon iconName={hasLiked ? 'HeartFill' : 'Heart'} />
            <span>{post.LikesCount || 0}</span>
          </button>
          <button
            className={styles.actionBtn}
            onClick={() => { this._toggleComments().catch(console.error); }}
            title="Comments"
          >
            <Icon iconName="Comment" />
            <span>{post.CommentsCount || 0}</span>
          </button>
        </div>

        {/* Expandable comments */}
        {showComments && (
          <CommentSection
            postId={post.Id}
            comments={comments}
            loading={commentsLoading}
            onAddComment={this._handleAddComment}
          />
        )}
      </div>
    );
  }
}
