import * as React from 'react';
import { ActionButton } from '@fluentui/react/lib/Button';
import importedStyles from './Announcements.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { IReactionSummary, ReactionType } from '../../../services/AnnouncementService';

interface IReactionBarProps {
  reactionSummary: IReactionSummary;
  onReactionClick: (reactionType: ReactionType) => void;
  loading?: boolean;
}

const ReactionBar: React.FC<IReactionBarProps> = ({ reactionSummary, onReactionClick, loading }) => {
  const reactions: Array<{ type: ReactionType; emoji: string; label: string }> = [
    { type: 'Like', emoji: '👍', label: 'Like' },
    { type: 'Love', emoji: '❤️', label: 'Love' },
    { type: 'Fire', emoji: '🔥', label: 'Important' }
  ];

  return (
    <div className={styles.reactionBar}>
      {reactions.map((reaction) => {
        const summary = reactionSummary[reaction.type];
        const isActive = summary.userReacted;
        const count = summary.count;

        return (
          <ActionButton
            key={reaction.type}
            className={`${styles.reactionButton} ${isActive ? styles.active : ''}`}
            onClick={() => onReactionClick(reaction.type)}
            disabled={loading}
            title={`${reaction.label}: ${count}`}
            ariaLabel={`${reaction.label} reaction, ${count} total`}
            text={`${reaction.emoji} ${reaction.label}${count > 0 ? ` (${count})` : ''}`}
          />
        );
      })}
    </div>
  );
};

export default ReactionBar;
