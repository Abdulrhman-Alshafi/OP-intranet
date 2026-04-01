import * as React from 'react';
import AnnouncementCard from './AnnouncementCard';
import { IAnnouncement } from './IAnnouncementsHubProps';
import styles from './AnnouncementsHub.module.scss';

export interface IAnnouncementGridProps {
  announcements: IAnnouncement[];
  showImages: boolean;
  enableAnimations: boolean;
  enableCategoryColors: boolean;
  accentColor: string;
  isAdmin: boolean;
  onClick?: (announcement: IAnnouncement) => void;
  onDismiss?: (id: number) => void;
  onDelete?: (id: number) => void;
}

const AnnouncementGrid: React.FC<IAnnouncementGridProps> = React.memo((props) => {
  const {
    announcements,
    showImages,
    enableAnimations,
    enableCategoryColors,
    accentColor,
    isAdmin,
    onClick,
    onDismiss,
    onDelete
  } = props;

  return (
    <div className={styles.grid} role="list" aria-label="Announcements grid">
      {announcements.map((a) => (
        <div key={a.Id} role="listitem">
          <AnnouncementCard
            announcement={a}
            showImages={showImages}
            enableAnimations={enableAnimations}
            enableCategoryColors={enableCategoryColors}
            accentColor={accentColor}
            isAdmin={isAdmin}
            onClick={onClick}
            onDismiss={onDismiss}
            onDelete={onDelete}
          />
        </div>
      ))}
    </div>
  );
});

AnnouncementGrid.displayName = 'AnnouncementGrid';

export default AnnouncementGrid;
