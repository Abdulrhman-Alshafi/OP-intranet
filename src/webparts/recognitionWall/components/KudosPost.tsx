import * as React from 'react';
import importedStyles from './RecognitionWall.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { IKudosPostProps } from './IRecognitionWallProps';
import { IComment } from '../../../services/RecognitionService';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import CommentSection from './CommentSection';

/**
 * KudosPost renders a single kudos card with sender→recipient, message,
 * category badge, like button, and expandable comments.
 */
const KudosPost = (props: IKudosPostProps): React.ReactElement => {
  const { post, hasLiked, isAdmin, onLoadComments, onToggleLike, onAddComment, currentUserId, onDelete } = props;
  const [showComments, setShowComments] = React.useState<boolean>(false);
  const [comments, setComments] = React.useState<IComment[]>([]);
  const [commentsLoading, setCommentsLoading] = React.useState<boolean>(false);
  const [liking, setLiking] = React.useState<boolean>(false);
  const [deleting, setDeleting] = React.useState<boolean>(false);

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
    return d.toLocaleDateString([], { year: 'numeric', month: 'short', day: 'numeric' });
  };

  const getCategoryClass = (): string => {
    switch (post.Category) {
      case 'Teamwork': return styles.catTeamwork;
      case 'Innovation': return styles.catInnovation;
      case 'Leadership': return styles.catLeadership;
      case 'Customer Focus': return styles.catCustomer;
      case 'Above & Beyond': return styles.catAboveBeyond;
      default: return styles.catDefault;
    }
  };

  const toggleComments = async (): Promise<void> => {
    const shouldShow = !showComments;
    setShowComments(shouldShow);

    if (shouldShow && comments.length === 0) {
      setCommentsLoading(true);
      try {
        const loadedComments = await onLoadComments(post.Id);
        setComments(loadedComments);
        setCommentsLoading(false);
      } catch {
        setCommentsLoading(false);
      }
    }
  };

  const handleLike = async (): Promise<void> => {
    if (liking) {
      return;
    }
    setLiking(true);
    try {
      await onToggleLike(post.Id);
    } finally {
      setLiking(false);
    }
  };

  const handleAddComment = (postId: number, text: string): void => {
    onAddComment(postId, text);
    const optimisticComment: IComment = {
      Id: Date.now(),
      PostId: postId,
      Comment: text,
      AuthorId: currentUserId,
      AuthorName: 'You',
      Created: new Date().toISOString()
    };
    setComments((currentComments) => [...currentComments, optimisticComment]);
  };

  const handleDelete = async (): Promise<void> => {
    if (deleting) {
      return;
    }
    setDeleting(true);
    try {
      await onDelete(post.Id);
    } catch {
      setDeleting(false);
    }
  };

  return (
    <div className={styles.kudosCard}>
      {isAdmin && (
        <IconButton
          iconProps={{ iconName: 'Cancel' }}
          title="Delete kudos"
          ariaLabel="Delete kudos"
          className={styles.deleteBtn}
          onClick={() => {
            handleDelete().catch(() => {
              setDeleting(false);
            });
          }}
          disabled={deleting}
        />
      )}
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

      {post.Category && (
        <span className={`${styles.categoryBadge} ${getCategoryClass()}`}>
          {post.Category}
        </span>
      )}

      <div className={styles.kudosMessage}>{post.Title}</div>

      <div className={styles.kudosDate}>
        <Icon iconName="Clock" className={styles.clockIcon} />
        <span>{formatDate(post.Created)}</span>
      </div>

      <div className={styles.kudosActions}>
        <button
          className={`${styles.actionBtn} ${hasLiked ? styles.liked : ''}`}
          onClick={() => {
            handleLike().catch(() => {
              setLiking(false);
            });
          }}
          disabled={liking}
          title={hasLiked ? 'Unlike' : 'Like'}
        >
          <Icon iconName={hasLiked ? 'HeartFill' : 'Heart'} />
          <span>{post.LikesCount || 0}</span>
        </button>
        <button
          className={styles.actionBtn}
          onClick={() => {
            toggleComments().catch(() => {
              setCommentsLoading(false);
            });
          }}
          title="Comments"
        >
          <Icon iconName="Comment" />
          <span>{post.CommentsCount || 0}</span>
        </button>
      </div>

      {showComments && (
        <CommentSection
          postId={post.Id}
          comments={comments}
          loading={commentsLoading}
          onAddComment={handleAddComment}
        />
      )}
    </div>
  );
};

export default KudosPost;
