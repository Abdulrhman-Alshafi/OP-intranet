import * as React from 'react';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { ActionButton } from '@fluentui/react/lib/Button';
import importedStyles from './Announcements.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { IAnnouncement, IComment, IReactionSummary, ReactionType } from '../../../services/AnnouncementService';
import ReactionBar from './ReactionBar';
import CommentSection from './CommentSection';

interface IAnnouncementCardProps {
  announcement: IAnnouncement;
  reactionSummary?: IReactionSummary;
  comments?: IComment[];
  enableReactions: boolean;
  enableComments: boolean;
  onReactionClick?: (reactionType: ReactionType) => void;
  onAddComment?: (text: string, parentId?: number) => Promise<void>;
  commentsLoading?: boolean;
  reactionsLoading?: boolean;
}

const AnnouncementCard: React.FC<IAnnouncementCardProps> = ({
  announcement,
  reactionSummary,
  comments,
  enableReactions,
  enableComments,
  onReactionClick,
  onAddComment,
  commentsLoading,
  reactionsLoading
}) => {
  const [expanded, setExpanded] = React.useState(false);

  const truncateContent = (content: string, lines: number = 3): string => {
    const lineArray = content.split('\n');
    if (lineArray.length > lines) {
      return lineArray.slice(0, lines).join('\n') + '...';
    }
    return content;
  };

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

  const displayContent = expanded ? announcement.Content : truncateContent(announcement.Content);
  const shouldShowExpand = announcement.Content.split('\n').length > 3;

  return (
    <div className={styles.announcementCard}>
      {/* Header */}
      <div className={styles.cardHeader}>
        <div className={styles.titleSection}>
          <h3>{announcement.Title}</h3>
        </div>
        <div className={styles.meta}>{formatDate(announcement.Created)}</div>
      </div>

      {/* Author */}
      <div className={styles.author}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <Persona
            text={announcement.Author?.Title}
            size={PersonaSize.size28}
            hidePersonaDetails
          />
          <div>
            <div className={styles.authorName}>{announcement.Author?.Title || 'Unknown'}</div>
            <div className={styles.authorEmail}>{announcement.Author?.EMail}</div>
          </div>
        </div>
      </div>

      {/* Content */}
      <div
        className={`${styles.cardContent} ${expanded ? styles.expanded : ''}`}
        // Content comes from SharePoint rich text field
        // eslint-disable-next-line react/no-danger
        dangerouslySetInnerHTML={{ __html: displayContent }}
      />

      {shouldShowExpand && (
        <ActionButton
          className={styles.expandButton}
          onClick={() => setExpanded(!expanded)}
          text={expanded ? 'Show less' : 'Show more'}
        />
      )}

      {/* Reactions */}
      {enableReactions && reactionSummary && onReactionClick && (
        <ReactionBar
          reactionSummary={reactionSummary}
          onReactionClick={onReactionClick}
          loading={reactionsLoading}
        />
      )}

      {/* Comments */}
      {enableComments && comments && onAddComment && (
        <CommentSection
          comments={comments}
          loading={commentsLoading}
          onAddComment={onAddComment}
        />
      )}
    </div>
  );
};

export default AnnouncementCard;
