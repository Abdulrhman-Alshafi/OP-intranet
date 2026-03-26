import * as React from 'react';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { ActionButton } from '@fluentui/react/lib/Button';
import importedStyles from './Announcements.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { IComment } from '../../../services/AnnouncementService';

interface IReplyItemProps {
  reply: IComment;
  onReplyClick: (parentId: number) => void;
}

const ReplyItem: React.FC<IReplyItemProps> = ({ reply, onReplyClick }) => {
  const [showNestedReplies, setShowNestedReplies] = React.useState<boolean>(true);

  const formatDate = (dateString: string): string => {
    try {
      const date = new Date(dateString);
      return date.toLocaleDateString(undefined, {
        month: 'short',
        day: 'numeric',
        year: date.getFullYear() !== new Date().getFullYear() ? 'numeric' : undefined
      });
    } catch {
      return dateString;
    }
  };

  const nestedCount = reply.Replies?.length || 0;

  return (
    <div className={`${styles.commentItem} ${styles.reply}`}>
      <div className={styles.commentHeader}>
        <div className={styles.authorInfoRow}>
          <Persona
            text={reply.Author?.Title || 'Unknown'}
            secondaryText={formatDate(reply.Created)}
            size={PersonaSize.size24}
            hidePersonaDetails={false}
          />
        </div>
      </div>

      <div className={styles.commentText}>{reply.Text}</div>

      <div className={styles.commentActions}>
        <ActionButton
          className={styles.actionButton}
          text="Reply"
          onClick={() => onReplyClick(reply.Id)}
        />
        {nestedCount > 0 && (
          <ActionButton
            className={styles.actionButton}
            text={showNestedReplies ? `Hide replies (${nestedCount})` : `Show replies (${nestedCount})`}
            onClick={() => setShowNestedReplies((prev) => !prev)}
          />
        )}
      </div>

      {nestedCount > 0 && (
        <div className={`${styles.repliesWrapper} ${showNestedReplies ? styles.expanded : styles.collapsed}`}>
          <div className={styles.repliesList}>
            {reply.Replies?.map((nestedReply) => (
              <ReplyItem
                key={nestedReply.Id}
                reply={nestedReply}
                onReplyClick={onReplyClick}
              />
            ))}
          </div>
        </div>
      )}
    </div>
  );
};

export default ReplyItem;
