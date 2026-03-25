import * as React from 'react';
import importedStyles from './RecognitionWall.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { ICommentSectionProps } from './IRecognitionWallProps';
import { IComment } from '../../../services/RecognitionService';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';

export interface ICommentSectionState {
  newComment: string;
  submitting: boolean;
}

/**
 * CommentSection renders comments for a single kudos post and a text input to add new ones.
 */
export default class CommentSection extends React.Component<ICommentSectionProps, ICommentSectionState> {
  constructor(props: ICommentSectionProps) {
    super(props);
    this.state = {
      newComment: '',
      submitting: false
    };
  }

  private _handleSubmit = async (): Promise<void> => {
    const text = this.state.newComment.trim();
    if (!text) return;

    this.setState({ submitting: true });
    try {
      this.props.onAddComment(this.props.postId, text);
      this.setState({ newComment: '', submitting: false });
    } catch {
      this.setState({ submitting: false });
    }
  };

  private _handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      this._handleSubmit().catch(console.error);
    }
  };

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
    return d.toLocaleDateString([], { month: 'short', day: 'numeric' });
  }

  public render(): React.ReactElement<ICommentSectionProps> {
    const { comments, loading } = this.props;
    const { newComment, submitting } = this.state;

    return (
      <div className={styles.commentSection}>
        {/* Comment list */}
        {loading && <Spinner size={SpinnerSize.small} label="Loading comments..." />}

        {!loading && comments.length === 0 && (
          <div className={styles.noComments}>No comments yet. Be the first!</div>
        )}

        {!loading && comments.map((comment: IComment) => (
          <div key={comment.Id} className={styles.commentItem}>
            <div className={styles.commentAuthor}>
              <Icon iconName="Contact" className={styles.commentIcon} />
              <span className={styles.commentAuthorName}>{comment.AuthorName}</span>
              <span className={styles.commentDate}>{this._formatDate(comment.Created)}</span>
            </div>
            <div className={styles.commentText}>{comment.Comment}</div>
          </div>
        ))}

        {/* Add comment input */}
        <div className={styles.commentInput}>
          <input
            type="text"
            className={styles.commentTextBox}
            placeholder="Write a comment..."
            value={newComment}
            onChange={(e) => this.setState({ newComment: e.target.value })}
            onKeyDown={this._handleKeyDown}
            disabled={submitting}
          />
          <button
            className={styles.commentSendBtn}
            onClick={() => { this._handleSubmit().catch(console.error); }}
            disabled={!newComment.trim() || submitting}
            title="Send comment"
          >
            <Icon iconName="Send" />
          </button>
        </div>
      </div>
    );
  }
}
