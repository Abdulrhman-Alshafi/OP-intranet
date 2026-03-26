import * as React from 'react';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { ActionButton } from '@fluentui/react/lib/Button';
import importedStyles from './Announcements.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { IComment } from '../../../services/AnnouncementService';
import ReplyItem from './ReplyItem';

interface ICommentItemProps {
  comment: IComment;
  onReplyClick: (parentId: number) => void;
}

const CommentItem: React.FC<ICommentItemProps> = ({ comment, onReplyClick }) => {
  const [showReplies, setShowReplies] = React.useState<boolean>(true);

  const formatDate = (dateString: string): string => {
    try {
      const date = new Date(dateString);
      const now = new Date();
      const diffMs = now.getTime() - date.getTime();
      const diffMins = Math.floor(diffMs / 60000);
      const diffHours = Math.floor(diffMs / 3600000);
      const diffDays = Math.floor(diffMs / 86400000);

      if (diffMins < 1) return 'just now';
      if (diffMins < 60) return `${diffMins}m ago`;
      if (diffHours < 24) return `${diffHours}h ago`;
      if (diffDays < 7) return `${diffDays}d ago`;

      return date.toLocaleDateString(undefined, { month: 'short', day: 'numeric' });
    } catch {
      return dateString;
    }
  };

  const replyCount = comment.Replies?.length || 0;

  return (
    <div className={styles.commentItem}>
      <div className={styles.commentHeader}>
        <div className={styles.authorInfoRow}>
          <Persona
            text={comment.Author?.Title || 'Unknown'}
            secondaryText={formatDate(comment.Created)}
            size={PersonaSize.size28}
            hidePersonaDetails={false}
          />
        </div>
      </div>

      <div className={styles.commentText}>{comment.Text}</div>

      <div className={styles.commentActions}>
        <ActionButton
          className={styles.actionButton}
          text="Reply"
          onClick={() => onReplyClick(comment.Id)}
        />
        {replyCount > 0 && (
          <ActionButton
            className={styles.actionButton}
            text={showReplies ? `Hide replies (${replyCount})` : `Show replies (${replyCount})`}
            onClick={() => setShowReplies((prev) => !prev)}
          />
        )}
      </div>

      {replyCount > 0 && (
        <div className={`${styles.repliesWrapper} ${showReplies ? styles.expanded : styles.collapsed}`}>
          <div className={styles.repliesList}>
            {comment.Replies?.map((reply) => (
              <ReplyItem
                key={reply.Id}
                reply={reply}
                onReplyClick={onReplyClick}
              />
            ))}
          </div>
        </div>
      )}
    </div>
  );
};

export default CommentItem;
