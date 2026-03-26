import * as React from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { ActionButton } from '@fluentui/react/lib/Button';
import importedStyles from './Announcements.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { IComment } from '../../../services/AnnouncementService';
import CommentItem from './CommentItem';

interface ICommentSectionProps {
  comments: IComment[];
  loading?: boolean;
  onAddComment: (text: string, parentId?: number) => Promise<void>;
}

const CommentSection: React.FC<ICommentSectionProps> = ({ comments, loading, onAddComment }) => {
  const [commentText, setCommentText] = React.useState<string>('');
  const [debouncedCommentText, setDebouncedCommentText] = React.useState<string>('');
  const [replyingTo, setReplyingTo] = React.useState<number | undefined>(undefined);
  const [submitting, setSubmitting] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string>('');
  const [visibleCount, setVisibleCount] = React.useState<number>(3);

  React.useEffect(() => {
    const handle = setTimeout(() => {
      setDebouncedCommentText(commentText);
    }, 250);

    return () => clearTimeout(handle);
  }, [commentText]);

  React.useEffect(() => {
    setVisibleCount(3);
  }, [comments]);

  const handleSubmitComment = async (): Promise<void> => {
    if (!debouncedCommentText.trim()) {
      return;
    }

    setSubmitting(true);
    setError('');

    try {
      await onAddComment(debouncedCommentText, replyingTo);
      setCommentText('');
      setDebouncedCommentText('');
      setReplyingTo(undefined);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Failed to add comment');
    } finally {
      setSubmitting(false);
    }
  };

  if (loading) {
    return (
      <div className={styles.commentSection}>
        <Spinner size={SpinnerSize.small} label="Loading comments..." />
      </div>
    );
  }

  return (
    <div className={styles.commentSection}>
      {/* Comment form */}
      <div className={styles.commentForm}>
        <TextField
          placeholder={replyingTo ? 'Write a reply...' : 'Add a comment...'}
          value={commentText}
          onChange={(_ev, newValue) => setCommentText(newValue || '')}
          disabled={submitting}
          multiline
          rows={2}
          className={styles.inputField}
        />
        <div style={{ display: 'flex', gap: 4 }}>
          <PrimaryButton
            iconProps={{ iconName: 'Send' }}
            onClick={handleSubmitComment}
            disabled={submitting || !debouncedCommentText.trim()}
            title={replyingTo ? 'Reply' : 'Comment'}
            ariaLabel={replyingTo ? 'Reply to comment' : 'Add comment'}
          />
          {replyingTo && (
            <DefaultButton
              iconProps={{ iconName: 'Cancel' }}
              onClick={() => {
                setReplyingTo(undefined);
                setCommentText('');
                setDebouncedCommentText('');
              }}
              disabled={submitting}
              title="Cancel reply"
              ariaLabel="Cancel reply"
            />
          )}
        </div>
      </div>

      {error && <div className={styles.errorMessage}>{error}</div>}

      {/* Comments list */}
      {comments.length > 0 ? (
        <div className={styles.commentsList}>
          {comments.slice(0, visibleCount).map((comment) => (
            <CommentItem
              key={comment.Id}
              comment={comment}
              onReplyClick={setReplyingTo}
            />
          ))}

          {visibleCount < comments.length && (
            <ActionButton
              className={styles.loadMoreButton}
              text={`Load more comments (${comments.length - visibleCount})`}
              onClick={() => setVisibleCount((prev) => prev + 3)}
            />
          )}
        </div>
      ) : (
        <p style={{ fontSize: '13px', color: 'var(--neutralSecondary)', margin: '12px 0 0 0' }}>
          No comments yet. Be the first to comment!
        </p>
      )}
    </div>
  );
};

export default CommentSection;
