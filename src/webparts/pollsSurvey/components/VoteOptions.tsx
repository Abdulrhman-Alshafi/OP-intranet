import * as React from 'react';
import styles from './PollsSurvey.module.scss';

export interface IVoteOptionsProps {
  /** Array of option strings to vote on */
  options: string[];
  /** Whether the user has already voted */
  hasVoted: boolean;
  /** Whether voting is in progress */
  isSubmitting: boolean;
  /** Whether the poll has expired */
  isExpired: boolean;
  /** Callback when user clicks an option to vote */
  onVote: (option: string) => void;
}

/**
 * Displays voting buttons for each poll option.
 * Disables all buttons if the user has already voted, the poll is expired, or a vote is being submitted.
 */
const VoteOptions: React.FC<IVoteOptionsProps> = ({
  options,
  hasVoted,
  isSubmitting,
  isExpired,
  onVote
}) => {
  const isDisabled = hasVoted || isSubmitting || isExpired;

  return (
    <div className={styles.voteOptions}>
      {options.map((option, index) => (
        <button
          key={index}
          className={`${styles.voteButton} ${isDisabled ? styles.voteButtonDisabled : ''}`}
          onClick={() => !isDisabled && onVote(option)}
          disabled={isDisabled}
          title={isExpired ? 'This poll has expired' : hasVoted ? 'You have already voted' : `Vote for: ${option}`}
        >
          <span className={styles.voteButtonIcon}>
            {/* Circle indicator */}
            <svg width="18" height="18" viewBox="0 0 18 18" fill="none">
              <circle cx="9" cy="9" r="8" stroke="currentColor" strokeWidth="1.5" />
            </svg>
          </span>
          <span className={styles.voteButtonText}>{option}</span>
        </button>
      ))}

      {isSubmitting && (
        <div className={styles.votingIndicator}>
          <div className={styles.spinner} />
          <span>Submitting your vote...</span>
        </div>
      )}

      {isExpired && !hasVoted && (
        <div className={styles.expiredNote}>This poll has expired. Voting is closed.</div>
      )}
    </div>
  );
};

export default VoteOptions;
