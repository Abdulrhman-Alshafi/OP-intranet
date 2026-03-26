import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Text } from '@fluentui/react/lib/Text';

import importedStyles from './Announcements.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;

import { IAnnouncementsProps } from './IAnnouncementsProps';
import {
  IAnnouncement,
  IComment,
  IReaction,
  ReactionType,
  AnnouncementService
} from '../../../services/AnnouncementService';

import CreateAnnouncementForm from './CreateAnnouncementForm';
import AnnouncementCard from './AnnouncementCard';

const Announcements: React.FC<IAnnouncementsProps> = (props) => {
  const {
    title,
    announcementsListName,
    enableCreateAnnouncement,
    enableReactions,
    enableComments,
    itemsPerPage,
    spHttpClient,
    siteUrl,
    userEmail,
    userId
  } = props;

  const reactionsListName = 'Reactions';
  const commentsListName = 'Comments';

  // ── State ─────────────────────────────────────────────────────────
  const [announcements, setAnnouncements] = useState<IAnnouncement[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | undefined>(undefined);
  const [currentPage, setCurrentPage] = useState<number>(1);

  // Reactions and comments per announcement
  const [reactionsMap, setReactionsMap] = useState<Map<number, IReaction[]>>(new Map());
  const [commentsMap, setCommentsMap] = useState<Map<number, IComment[]>>(new Map());
  const [loadingReactionsId, setLoadingReactionsId] = useState<number | undefined>(undefined);
  const [loadingCommentsId, setLoadingCommentsId] = useState<number | undefined>(undefined);

  const loadAnnouncements = async (): Promise<void> => {
    setLoading(true);
    setError(undefined);

    try {
      const data = await AnnouncementService.getAnnouncements(
        siteUrl,
        announcementsListName,
        spHttpClient
      );
      setAnnouncements(data);
    } catch (err) {
      setError(
        err instanceof Error
          ? err.message
          : 'Failed to load announcements. Please check the list name in settings.'
      );
    } finally {
      setLoading(false);
    }
  };

  // ── Data fetching ─────────────────────────────────────────────────
  useEffect(() => {
    if (announcementsListName) {
      loadAnnouncements().catch(console.error);
    }
  }, [announcementsListName, siteUrl, spHttpClient]);

  // ── Fetch reactions for an announcement ────────────────────────────
  const fetchReactions = async (announcementId: number): Promise<void> => {
    if (!enableReactions) return;

    setLoadingReactionsId(announcementId);
    try {
      const reactions = await AnnouncementService.getReactions(
        siteUrl,
        reactionsListName,
        announcementId,
        spHttpClient
      );
      setReactionsMap((prev) => new Map(prev).set(announcementId, reactions));
    } catch (err) {
      console.error('Failed to fetch reactions:', err);
    } finally {
      setLoadingReactionsId(undefined);
    }
  };

  // ── Fetch comments for an announcement ─────────────────────────────
  const fetchComments = async (announcementId: number): Promise<void> => {
    if (!enableComments) return;

    setLoadingCommentsId(announcementId);
    try {
      const comments = await AnnouncementService.getComments(
        siteUrl,
        commentsListName,
        announcementId,
        spHttpClient
      );
      setCommentsMap((prev) => new Map(prev).set(announcementId, comments));
    } catch (err) {
      console.error('Failed to fetch comments:', err);
    } finally {
      setLoadingCommentsId(undefined);
    }
  };

  // ── Handle reaction toggle ────────────────────────────────────────
  const handleReactionToggle = async (announcementId: number, reactionType: ReactionType): Promise<void> => {
    const reactions = reactionsMap.get(announcementId) || [];
    const existingReaction = reactions.find(
      (r) => r.ReactionType === reactionType && (r.UserId === userId || r.UserEmail === userEmail)
    );
    const previous = [...reactions];

    const optimistic: IReaction[] = existingReaction
      ? reactions.filter((r) => r.Id !== existingReaction.Id)
      : [
          ...reactions,
          {
            Id: -Date.now(),
            AnnouncementId: announcementId,
            ReactionType: reactionType,
            UserId: userId,
            UserEmail: userEmail,
            Created: new Date().toISOString()
          }
        ];

    setReactionsMap((prev) => new Map(prev).set(announcementId, optimistic));

    try {
      setLoadingReactionsId(announcementId);
      await AnnouncementService.setReaction(
        siteUrl,
        reactionsListName,
        announcementId,
        reactionType,
        userId,
        userEmail,
        spHttpClient,
        existingReaction?.Id
      );

      // Refresh reactions
      await fetchReactions(announcementId);
    } catch (err) {
      console.error('Failed to update reaction:', err);
      setReactionsMap((prev) => new Map(prev).set(announcementId, previous));
    } finally {
      setLoadingReactionsId(undefined);
    }
  };

  // ── Handle add comment ────────────────────────────────────────────
  const handleAddComment = async (
    announcementId: number,
    text: string,
    parentCommentId?: number
  ): Promise<void> => {
    try {
      await AnnouncementService.createComment(
        siteUrl,
        commentsListName,
        announcementId,
        text,
        spHttpClient,
        parentCommentId
      );

      // Refresh comments
      await fetchComments(announcementId);
    } catch (err) {
      console.error('Failed to add comment:', err);
      throw err;
    }
  };

  // ── Handle create announcement ────────────────────────────────────
  const handleCreateAnnouncement = async (announcementTitle: string, announcementContent: string): Promise<void> => {
    try {
      await AnnouncementService.createAnnouncement(
        siteUrl,
        announcementsListName,
        announcementTitle,
        announcementContent,
        spHttpClient
      );

      await loadAnnouncements();
      setCurrentPage(1);
    } catch (err) {
      console.error('Failed to create announcement:', err);
      throw err;
    }
  };

  // ── Pagination ────────────────────────────────────────────────────
  const totalPages = useMemo(
    () => Math.max(1, Math.ceil(announcements.length / itemsPerPage)),
    [announcements.length, itemsPerPage]
  );

  const visibleAnnouncements = useMemo(() => {
    const startIndex = (currentPage - 1) * itemsPerPage;
    return announcements.slice(startIndex, startIndex + itemsPerPage);
  }, [announcements, currentPage, itemsPerPage]);

  // ── Lazy load reactions and comments on first render ───────────────
  useEffect(() => {
    visibleAnnouncements.forEach((ann) => {
      if (enableReactions && !reactionsMap.has(ann.Id)) {
        fetchReactions(ann.Id).catch(console.error);
      }
      if (enableComments && !commentsMap.has(ann.Id)) {
        fetchComments(ann.Id).catch(console.error);
      }
    });
  }, [visibleAnnouncements, enableReactions, enableComments]);

  // ── Render ────────────────────────────────────────────────────────
  if (loading) {
    return (
      <div className={`${styles.announcements} ${styles.loadingContainer}`}>
        <Spinner size={SpinnerSize.large} label="Loading announcements..." />
      </div>
    );
  }

  return (
    <div className={styles.announcements}>
      {/* Header */}
      <div className={styles.header}>
        <Text as="h2" block style={{ margin: '0 0 4px 0' }}>
          {title}
        </Text>
        <span className={styles.summary}>
          {announcements.length} announcement{announcements.length !== 1 ? 's' : ''}
        </span>
      </div>

      {/* Error */}
      {error && (
        <MessageBar
          messageBarType={MessageBarType.error}
          isMultiline={false}
          onDismiss={() => setError(undefined)}
          dismissButtonAriaLabel="Close"
          styles={{ root: { marginBottom: 16 } }}
        >
          {error}
        </MessageBar>
      )}

      {/* Create announcement form */}
      {enableCreateAnnouncement && !error && (
        <CreateAnnouncementForm
          onSubmit={handleCreateAnnouncement}
          loading={loading}
        />
      )}

      {/* Announcements list */}
      {visibleAnnouncements.length > 0 ? (
        <div className={styles.announcementsList}>
          {visibleAnnouncements.map((announcement) => {
            const reactions = reactionsMap.get(announcement.Id) || [];
            const comments = commentsMap.get(announcement.Id) || [];
            const reactionSummary = enableReactions
              ? AnnouncementService.buildReactionSummary(reactions, userEmail, userId)
              : undefined;

            return (
              <AnnouncementCard
                key={announcement.Id}
                announcement={announcement}
                reactionSummary={reactionSummary}
                comments={comments}
                enableReactions={enableReactions}
                enableComments={enableComments}
                onReactionClick={(reactionType) => handleReactionToggle(announcement.Id, reactionType)}
                onAddComment={(text, parentId) => handleAddComment(announcement.Id, text, parentId)}
                reactionsLoading={loadingReactionsId === announcement.Id}
                commentsLoading={loadingCommentsId === announcement.Id}
              />
            );
          })}
        </div>
      ) : (
        <div className={styles.emptyState}>
          <p>No announcements yet. Stay tuned!</p>
        </div>
      )}

      {/* Pagination */}
      {totalPages > 1 && (
        <div className={styles.pagination}>
          <DefaultButton
            text="Previous"
            iconProps={{ iconName: 'ChevronLeft' }}
            disabled={currentPage === 1}
            onClick={() => setCurrentPage((p) => p - 1)}
          />
          <Text className={styles.pageInfo}>
            Page {currentPage} of {totalPages}
          </Text>
          <DefaultButton
            text="Next"
            iconProps={{ iconName: 'ChevronRight' }}
            disabled={currentPage === totalPages}
            onClick={() => setCurrentPage((p) => p + 1)}
          />
        </div>
      )}
    </div>
  );
};

export default Announcements;
