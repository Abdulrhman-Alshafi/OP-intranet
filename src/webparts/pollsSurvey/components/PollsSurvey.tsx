import * as React from 'react';
import styles from './PollsSurvey.module.scss';
import { IPollsSurveyProps } from './IPollsSurveyProps';
import { PollService, IPoll, IPollResults } from '../../../services/PollService';
import PollCard from './PollCard';
import CreatePollForm from './CreatePollForm';
import { PrimaryButton, IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { Toggle } from '@fluentui/react/lib/Toggle';

/**
 * Main container component for Polls & Quick Surveys.
 *
 * Responsibilities:
 * - Fetches active polls from SharePoint on mount
 * - Batch-loads results for all visible polls
 * - Provides "Create Poll" form for admins
 * - Handles loading, error, and empty states
 */
const PollsSurvey: React.FC<IPollsSurveyProps> = (props) => {
  const {
    title,
    refreshInterval,
    allowAnonymous,
    pollsPerPage,
    isAdmin,
    autoHideDays,
    spHttpClient,
    siteUrl,
    currentUserId
  } = props;

  // Create a stable PollService instance
  const pollServiceRef = React.useRef<PollService | null>(null);
  if (!pollServiceRef.current) {
    pollServiceRef.current = new PollService(spHttpClient, siteUrl);
  }
  const pollService = pollServiceRef.current;

  const [polls, setPolls] = React.useState<IPoll[]>([]);
  const [resultsMap, setResultsMap] = React.useState<{ [pollId: number]: IPollResults }>({});
  const [isLoading, setIsLoading] = React.useState(true);
  const [error, setError] = React.useState('');
  const [showCreateForm, setShowCreateForm] = React.useState(false);
  const [isCreating, setIsCreating] = React.useState(false);
  const [visibleCount, setVisibleCount] = React.useState(pollsPerPage);
  const [showHidden, setShowHidden] = React.useState(false);

  // ── Data Fetching ──────────────────────────────────────────────────
  const loadPolls = React.useCallback(async (showLoading: boolean = false): Promise<void> => {
    if (showLoading) setIsLoading(true);
    setError('');

    try {
      const fetchedPolls = await pollService.getPolls();
      setPolls(fetchedPolls);

      // Batch-fetch results for all polls at once (optimization)
      if (fetchedPolls.length > 0) {
        const batchResults = await pollService.getBatchResults(fetchedPolls, currentUserId);
        setResultsMap(batchResults);
      }
    } catch (err) {
      console.error('PollsSurvey: Failed to load polls', err);
      setError(err instanceof Error ? err.message : 'Failed to load polls.');
    } finally {
      setIsLoading(false);
    }
  }, [pollService, currentUserId]);

  // Initial load
  React.useEffect(() => {
    loadPolls(true).catch(console.error);
  }, [loadPolls]);

  // Auto-refresh removed per user request

  // ── Create Poll Handler ────────────────────────────────────────────
  const handleCreatePoll = async (
    question: string,
    choices: string[],
    closingDate: string | null
  ): Promise<void> => {
    setIsCreating(true);
    try {
      await pollService.createPoll(question, choices, closingDate);
      setShowCreateForm(false);
      await loadPolls(false);
    } catch (err) {
      console.error('PollsSurvey: Failed to create poll', err);
    } finally {
      setIsCreating(false);
    }
  };

  // ── Manual refresh ─────────────────────────────────────────────────
  const handleRefresh = (): void => {
    loadPolls(false).catch(console.error);
  };

  // ── Show More Polls ────────────────────────────────────────────────
  const handleShowMore = (): void => {
    setVisibleCount(prev => prev + pollsPerPage);
  };

  // ── Client-side visibility filtering ─────────────────────────────────
  // 1. Pinned polls always show (immune to hide + auto-hide)
  // 2. Hidden polls are excluded unless admin toggles "Show Hidden"
  // 3. Expired polls older than autoHideDays are auto-hidden (unless pinned)
  const filteredPolls = React.useMemo((): IPoll[] => {
    const now = new Date();
    const autoHideMs = autoHideDays * 24 * 60 * 60 * 1000;

    return polls.filter((poll) => {
      // Pinned always shows
      if (poll.IsPinned) return true;

      // Hidden: only show when admin has toggled "Show Hidden"
      if (poll.IsHidden) return showHidden && isAdmin;

      // Auto-hide: expired polls older than autoHideDays
      if (!poll.IsActive && poll.ClosingDate) {
        const closingDate = new Date(poll.ClosingDate);
        if (now.getTime() - closingDate.getTime() > autoHideMs) {
          // Auto-hidden — only show when admin toggles "Show Hidden"
          return showHidden && isAdmin;
        }
      }

      return true;
    });
  }, [polls, autoHideDays, showHidden, isAdmin]);

  // Sort: pinned polls first, then by created date (already sorted desc from API)
  const sortedPolls = React.useMemo((): IPoll[] => {
    const pinned = filteredPolls.filter(p => p.IsPinned);
    const unpinned = filteredPolls.filter(p => !p.IsPinned);
    return pinned.concat(unpinned);
  }, [filteredPolls]);

  const visiblePolls = sortedPolls.slice(0, visibleCount);
  const hasMore = sortedPolls.length > visibleCount;

  return (
    <div className={styles.pollsSurvey}>
      {/* Header */}
      <div className={styles.header}>
        <div className={styles.titleSection}>
          <h2>{title}</h2>
          {!isLoading && (
            <div className={styles.summary}>
              <div className={styles.summaryItem}>
                <Icon iconName="SurveyQuestions" className={styles.icon} />
                <span>{sortedPolls.length} {sortedPolls.length === 1 ? 'poll' : 'polls'}</span>
              </div>
            </div>
          )}
        </div>
        <div className={styles.controls}>
          {isAdmin && (
            <PrimaryButton
              text="New Poll"
              iconProps={{ iconName: 'Add' }}
              onClick={() => setShowCreateForm(true)}
            />
          )}
          <IconButton
            iconProps={{ iconName: 'Refresh' }}
            title="Refresh polls"
            ariaLabel="Refresh polls"
            onClick={handleRefresh}
            disabled={isLoading}
          />
        </div>
      </div>

      {/* Admin: Show Hidden toggle */}
      {isAdmin && (
        <div className={styles.showHiddenToggle}>
          <Toggle
            label="Show hidden polls"
            inlineLabel
            checked={showHidden}
            onChange={(_ev, checked) => setShowHidden(!!checked)}
            styles={{ root: { marginBottom: 0 } }}
          />
        </div>
      )}

      {/* Loading State */}
      {isLoading && (
        <div className={styles.loadingContainer}>
          <div className={styles.spinner} />
          <p>Loading polls...</p>
        </div>
      )}

      {/* Error State */}
      {error && !isLoading && (
        <div className={styles.errorContainer}>
          <svg width="24" height="24" viewBox="0 0 24 24" fill="none">
            <circle cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="1.5" />
            <path d="M12 8V12M12 16H12.01" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
          </svg>
          <p>{error}</p>
          <button className={styles.secondaryButton} onClick={handleRefresh}>
            Try Again
          </button>
        </div>
      )}

      {/* Empty State */}
      {!isLoading && !error && sortedPolls.length === 0 && (
        <div className={styles.emptyContainer}>
          <svg width="48" height="48" viewBox="0 0 48 48" fill="none">
            <rect x="8" y="10" width="32" height="28" rx="4" stroke="currentColor" strokeWidth="2" />
            <path d="M16 22H32M16 28H28M16 34H24" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
            <circle cx="36" cy="12" r="8" fill="#0078d4" stroke="#fff" strokeWidth="2" />
            <path d="M36 9V15M33 12H39" stroke="#fff" strokeWidth="1.5" strokeLinecap="round" />
          </svg>
          <p>No active polls yet.</p>
          {isAdmin && (
            <button className={styles.primaryButton} onClick={() => setShowCreateForm(true)}>
              Create Your First Poll
            </button>
          )}
        </div>
      )}

      {/* Polls Grid */}
      {!isLoading && !error && sortedPolls.length > 0 && (
        <div className={styles.pollsGrid}>
          {visiblePolls.map((poll) => (
            <PollCard
              key={poll.Id}
              poll={poll}
              allowAnonymous={allowAnonymous}
              currentUserId={currentUserId}
              pollService={pollService}
              initialResults={resultsMap[poll.Id] || null}
              isAdmin={isAdmin}
              onVoteSubmitted={handleRefresh}
              onPollUpdated={handleRefresh}
            />
          ))}
        </div>
      )}

      {/* Show More */}
      {hasMore && (
        <div className={styles.showMoreContainer}>
          <button className={styles.secondaryButton} onClick={handleShowMore}>
            Show More Polls ({sortedPolls.length - visibleCount} remaining)
          </button>
        </div>
      )}

      {/* Create Poll Form (Modal overlay) */}
      <CreatePollForm
        isOpen={showCreateForm}
        isSubmitting={isCreating}
        onCreate={handleCreatePoll}
        onClose={() => setShowCreateForm(false)}
      />
    </div>
  );
};

export default PollsSurvey;
