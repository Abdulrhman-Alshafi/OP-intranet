import * as React from 'react';
import styles from './PollsSurvey.module.scss';
import { IPoll, IPollResults, PollService } from '../../../services/PollService';
import VoteOptions from './VoteOptions';
import ResultsChart from './ResultsChart';
import { IconButton } from '@fluentui/react/lib/Button';

export interface IPollCardProps {
  /** The poll data */
  poll: IPoll;
  /** Whether anonymous voting is allowed */
  allowAnonymous: boolean;
  /** Current user's SharePoint numeric ID */
  currentUserId: number;
  /** PollService instance */
  pollService: PollService;
  /** Pre-loaded results (from batch fetch) */
  initialResults: IPollResults | undefined;
  /** Whether the current user is a poll admin */
  isAdmin: boolean;
  /** Callback after a vote is successfully submitted */
  onVoteSubmitted: () => void;
  /** Callback after poll admin action (pin, hide, delete) */
  onPollUpdated: () => void;
}

/**
 * PollCard displays a single poll with its question, countdown timer,
 * voting buttons or results chart (depending on whether the user has voted).
 *
 * State transitions:
 * - Loading → Vote view (if user hasn't voted) or Results view (if voted/expired)
 * - Vote view → Submitting → Results view (after successful vote)
 */
const PollCard: React.FC<IPollCardProps> = ({
  poll,
  allowAnonymous,
  currentUserId,
  pollService,
  initialResults,
  isAdmin,
  onVoteSubmitted,
  onPollUpdated
}) => {
  const [results, setResults] = React.useState<IPollResults | null>(initialResults ?? null);
  const [isSubmitting, setIsSubmitting] = React.useState(false);
  const [voteError, setVoteError] = React.useState('');
  const [countdown, setCountdown] = React.useState('');
  const [isExpired, setIsExpired] = React.useState(!poll.IsActive);

  // Parse choices from JSON string (uses updated field name PollChoices)
  let choices: string[] = [];
  try {
    choices = JSON.parse(poll.PollChoices);
  } catch { choices = []; }

  const hasVoted = results?.currentUserVote !== null && results?.currentUserVote !== undefined;

  // Update results when initialResults changes (from parent refresh)
  React.useEffect(() => {
    if (initialResults) {
      setResults(initialResults);
    }
  }, [initialResults]);

  // ── Countdown Timer (uses ClosingDate field) ───────────────────────
  React.useEffect(() => {
    if (!poll.ClosingDate) {
      setCountdown('');
      return;
    }

    const updateCountdown = (): void => {
      const now = new Date().getTime();
      const closing = new Date(poll.ClosingDate!).getTime();
      const diff = closing - now;

      if (diff <= 0) {
        setCountdown('Expired');
        setIsExpired(true);
        return;
      }

      const days = Math.floor(diff / (1000 * 60 * 60 * 24));
      const hours = Math.floor((diff % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
      const minutes = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));
      const seconds = Math.floor((diff % (1000 * 60)) / 1000);

      if (days > 0) {
        setCountdown(`${days}d ${hours}h ${minutes}m`);
      } else if (hours > 0) {
        setCountdown(`${hours}h ${minutes}m ${seconds}s`);
      } else {
        setCountdown(`${minutes}m ${seconds}s`);
      }
    };

    updateCountdown();
    const interval = setInterval(updateCountdown, 1000);
    return () => clearInterval(interval);
  }, [poll.ClosingDate]);

  // ── Vote Handler (uses currentUserId for Person field) ─────────────
  const handleVote = async (selectedChoice: string): Promise<void> => {
    setIsSubmitting(true);
    setVoteError('');

    try {
      await pollService.submitVote(
        poll.Id,
        selectedChoice,
        currentUserId
      );

      // Refresh results after voting
      const updatedResults = await pollService.getPollResults(
        poll.Id,
        choices,
        currentUserId
      );
      setResults(updatedResults);
      onVoteSubmitted();
    } catch (err) {
      setVoteError(err instanceof Error ? err.message : 'Failed to submit vote.');
    } finally {
      setIsSubmitting(false);
    }
  };

  // ── Export Handler ─────────────────────────────────────────────────
  const handleExport = (): void => {
    if (!results) return;
    const csv = pollService.generateResultsCsv(poll, results);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `poll-results-${poll.Id}.csv`;
    link.click();
    URL.revokeObjectURL(url);
  };

  // ── Admin Action Handlers ─────────────────────────────────────────
  const handleTogglePin = async (): Promise<void> => {
    try {
      await pollService.togglePin(poll.Id, poll.IsPinned);
      onPollUpdated();
    } catch (err) {
      console.error('Failed to toggle pin', err);
    }
  };

  const handleToggleHidden = async (): Promise<void> => {
    try {
      await pollService.toggleHidden(poll.Id, poll.IsHidden);
      onPollUpdated();
    } catch (err) {
      console.error('Failed to toggle hidden', err);
    }
  };

  const handleDelete = async (): Promise<void> => {
    // eslint-disable-next-line no-restricted-globals
    const confirmed = confirm(`Delete poll "${poll.Title}" and all its votes? This cannot be undone.`);
    if (!confirmed) return;
    try {
      await pollService.deletePoll(poll.Id);
      onPollUpdated();
    } catch (err) {
      console.error('Failed to delete poll', err);
    }
  };

  // ── Format creation date ──────────────────────────────────────────
  const createdDate = new Date(poll.Created).toLocaleDateString(undefined, {
    year: 'numeric', month: 'short', day: 'numeric'
  });

  return (
    <div className={`${styles.pollCard} ${isExpired ? styles.pollCardExpired : ''} ${poll.IsHidden ? styles.pollCardHidden : ''}`}>
      {/* Card Header */}
      <div className={styles.pollCardHeader}>
        <div className={styles.pollHeaderTop}>
          <div className={styles.pollQuestion}>
            {poll.IsPinned && <span className={styles.pinnedBadge} title="Pinned">📌</span>}
            {poll.IsHidden && <span className={styles.hiddenBadge} title="Hidden">👁️‍🗨️</span>}
            {poll.Title}
          </div>

          {/* Admin action buttons */}
          {isAdmin && (
            <div className={styles.adminActions}>
              <IconButton
                iconProps={{ iconName: poll.IsPinned ? 'Unpin' : 'Pin' }}
                title={poll.IsPinned ? 'Unpin poll' : 'Pin poll'}
                ariaLabel={poll.IsPinned ? 'Unpin poll' : 'Pin poll'}
                onClick={handleTogglePin}
                className={styles.adminActionButton}
              />
              <IconButton
                iconProps={{ iconName: poll.IsHidden ? 'View' : 'Hide3' }}
                title={poll.IsHidden ? 'Show poll' : 'Hide poll'}
                ariaLabel={poll.IsHidden ? 'Show poll' : 'Hide poll'}
                onClick={handleToggleHidden}
                className={styles.adminActionButton}
              />
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                title="Delete poll"
                ariaLabel="Delete poll"
                onClick={handleDelete}
                className={styles.adminDeleteButton}
              />
            </div>
          )}
        </div>

        <div className={styles.pollMeta}>
          <span className={styles.pollAuthor}>
            <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
              <circle cx="7" cy="5" r="3" stroke="currentColor" strokeWidth="1.2" />
              <path d="M2 13C2 10.2 4.2 8 7 8C9.8 8 12 10.2 12 13" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" />
            </svg>
            {poll.CreatedBy}
          </span>
          <span className={styles.pollDate}>{createdDate}</span>
        </div>

        {/* Countdown badge (ClosingDate) */}
        {poll.ClosingDate && (
          <div className={`${styles.countdownBadge} ${isExpired ? styles.countdownExpired : ''}`}>
            <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
              <circle cx="7" cy="7" r="6" stroke="currentColor" strokeWidth="1.2" />
              <path d="M7 4V7L9 9" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
            {isExpired ? 'Expired' : countdown}
          </div>
        )}
      </div>

      {/* Card Body: Vote or Results */}
      <div className={styles.pollCardBody}>
        {hasVoted || isExpired ? (
          <ResultsChart
            options={choices}
            results={results!}
            onExport={handleExport}
          />
        ) : (
          <VoteOptions
            options={choices}
            hasVoted={hasVoted}
            isSubmitting={isSubmitting}
            isExpired={isExpired}
            onVote={handleVote}
          />
        )}

        {voteError && (
          <div className={styles.errorMessage}>{voteError}</div>
        )}
      </div>
    </div>
  );
};

export default PollCard;
