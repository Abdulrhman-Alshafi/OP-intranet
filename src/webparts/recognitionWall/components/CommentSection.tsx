import * as React from 'react';
import importedStyles from './RecognitionWall.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { ICommentSectionProps } from './IRecognitionWallProps';
import { IComment } from '../../../services/RecognitionService';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';

/**
 * CommentSection renders comments for a single kudos post and a text input to add new ones.
 */
const CommentSection = (props: ICommentSectionProps): React.ReactElement => {
  const { comments, loading, onAddComment, postId } = props;
  const [newComment, setNewComment] = React.useState<string>('');
  const [submitting, setSubmitting] = React.useState<boolean>(false);

  const formatDate = (dateStr: string): string => {
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
  };

  const handleSubmit = async (): Promise<void> => {
    const text = newComment.trim();
    if (!text) {
      return;
    }

    setSubmitting(true);
    try {
      onAddComment(postId, text);
      setNewComment('');
      setSubmitting(false);
    } catch {
      setSubmitting(false);
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>): void => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      handleSubmit().catch(() => {
        setSubmitting(false);
      });
    }
  };

  return (
    <div className={styles.commentSection}>
      {loading && <Spinner size={SpinnerSize.small} label="Loading comments..." />}

      {!loading && comments.length === 0 && (
        <div className={styles.noComments}>No comments yet. Be the first!</div>
      )}

      {!loading && comments.map((comment: IComment) => (
        <div key={comment.Id} className={styles.commentItem}>
          <div className={styles.commentAuthor}>
            <Icon iconName="Contact" className={styles.commentIcon} />
            <span className={styles.commentAuthorName}>{comment.AuthorName}</span>
            <span className={styles.commentDate}>{formatDate(comment.Created)}</span>
          </div>
          <div className={styles.commentText}>{comment.Comment}</div>
        </div>
      ))}

      <div className={styles.commentInput}>
        <input
          type="text"
          className={styles.commentTextBox}
          placeholder="Write a comment..."
          value={newComment}
          onChange={(e) => setNewComment(e.target.value)}
          onKeyDown={handleKeyDown}
          disabled={submitting}
        />
        <button
          className={styles.commentSendBtn}
          onClick={() => {
            handleSubmit().catch(() => {
              setSubmitting(false);
            });
          }}
          disabled={!newComment.trim() || submitting}
          title="Send comment"
        >
          <Icon iconName="Send" />
        </button>
      </div>
    </div>
  );
};

export default CommentSection;
