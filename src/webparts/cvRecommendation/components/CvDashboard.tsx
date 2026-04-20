import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { ICvStats } from '../../../services/CvRecommendationService';
import styles from './CvRecommendation.module.scss';

export interface ICvDashboardProps {
  stats: ICvStats;
  accentColor: string;
}

const CvDashboard: React.FC<ICvDashboardProps> = ({ stats }) => {
  const cards = [
    { label: 'Total CVs', value: stats.total, icon: 'People', cls: styles.statTotal },
    { label: 'Submitted', value: stats.submitted, icon: 'Mail', cls: styles.statSubmitted },
    { label: 'Under Review', value: stats.underReview, icon: 'Search', cls: styles.statReview },
    { label: 'Accepted', value: stats.accepted, icon: 'CheckMark', cls: styles.statAccepted },
    { label: 'Rejected', value: stats.rejected, icon: 'Cancel', cls: styles.statRejected }
  ];

  return (
    <div className={styles.statsGrid}>
      {cards.map(card => (
        <div key={card.label} className={`${styles.statCard} ${card.cls}`}>
          <Icon iconName={card.icon} style={{ fontSize: 18, marginBottom: 4 }} />
          <span className={styles.statValue}>{card.value}</span>
          <span className={styles.statLabel}>{card.label}</span>
        </div>
      ))}
    </div>
  );
};

export default CvDashboard;
