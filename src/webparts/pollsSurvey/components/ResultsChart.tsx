import * as React from 'react';
import styles from './PollsSurvey.module.scss';
import { IPollResults } from '../../../services/PollService';

// Color palette for options (vibrant, modern)
const BAR_COLORS = [
  '#0078d4', // Blue
  '#00b294', // Teal
  '#8764b8', // Purple
  '#e3008c', // Magenta
  '#ff8c00', // Orange
  '#107c10', // Green
];

export interface IResultsChartProps {
  /** Poll options array */
  options: string[];
  /** Aggregated results data */
  results: IPollResults;
  /** Allow exporting results to CSV */
  onExport?: () => void;
}

/**
 * Displays poll results as a horizontal bar chart.
 * - Shows percentage + vote count per option
 * - Highlights the option the current user voted for
 * - Shows total vote count at the bottom
 * - Smooth animated bars
 */
const ResultsChart: React.FC<IResultsChartProps> = ({ options, results, onExport }) => {
  return (
    <div className={styles.resultsChart}>
      <div className={styles.resultsBars}>
        {options.map((option, index) => {
          const count = results.optionCounts[option] || 0;
          const percentage = results.optionPercentages[option] || 0;
          const isUserVote = results.currentUserVote === option;
          const barColor = BAR_COLORS[index % BAR_COLORS.length];

          return (
            <div
              key={index}
              className={`${styles.resultRow} ${isUserVote ? styles.resultRowHighlighted : ''}`}
            >
              <div className={styles.resultLabel}>
                <span className={styles.resultOptionText}>{option}</span>
                {isUserVote && (
                  <span className={styles.yourVoteBadge} title="Your vote">
                    <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
                      <circle cx="7" cy="7" r="6" fill={barColor} />
                      <path d="M4 7L6 9L10 5" stroke="#fff" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                  </span>
                )}
              </div>
              <div className={styles.barContainer}>
                <div
                  className={styles.bar}
                  style={{
                    width: `${Math.max(percentage, 2)}%`,
                    backgroundColor: barColor,
                  }}
                />
                <span className={styles.barStats}>
                  {percentage}% ({count} {count === 1 ? 'vote' : 'votes'})
                </span>
              </div>
            </div>
          );
        })}
      </div>

      <div className={styles.resultsFooter}>
        <span className={styles.totalVotes}>
          Total: <strong>{results.totalVotes}</strong> {results.totalVotes === 1 ? 'vote' : 'votes'}
        </span>
        {onExport && (
          <button className={styles.exportButton} onClick={onExport} title="Export results to CSV">
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
              <path d="M8 2V10M8 10L5 7M8 10L11 7" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
              <path d="M2 12H14" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round"/>
            </svg>
            Export
          </button>
        )}
      </div>
    </div>
  );
};

export default ResultsChart;
