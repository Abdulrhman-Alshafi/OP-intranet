import * as React from 'react';
import importedStyles from './TaskDashboard.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import { ITaskCardProps } from './ITaskDashboardProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Icon } from '@fluentui/react/lib/Icon';

/**
 * TaskCard component displays a single task with its details
 */
const TaskCard = (props: ITaskCardProps): React.ReactElement => {
  const { task, showSource, highlightOverdue } = props;

  const isOverdue = (): boolean => {
    if (!task.dueDate || task.status === 'Completed') {
      return false;
    }
    return new Date(task.dueDate) < new Date();
  };

  const isDueToday = (): boolean => {
    if (!task.dueDate) {
      return false;
    }

    const today = new Date();
    const dueDate = new Date(task.dueDate);

    return (
      today.getDate() === dueDate.getDate() &&
      today.getMonth() === dueDate.getMonth() &&
      today.getFullYear() === dueDate.getFullYear()
    );
  };

  const formatDate = (date: Date | null): string => {
    if (!date) {
      return 'No due date';
    }

    const dateObj = new Date(date);
    const now = new Date();
    const diffTime = dateObj.getTime() - now.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

    if (isDueToday()) {
      return 'Due Today';
    }
    if (diffDays === 1) {
      return 'Due Tomorrow';
    }
    if (diffDays === -1) {
      return 'Due Yesterday';
    }
    if (diffDays < -1) {
      return `${Math.abs(diffDays)} days overdue`;
    }
    if (diffDays > 0 && diffDays <= 7) {
      return `Due in ${diffDays} days`;
    }

    return dateObj.toLocaleDateString([], {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    });
  };

  const getStatusClass = (): string => {
    switch (task.status) {
      case 'Not Started':
        return styles.notStarted;
      case 'In Progress':
        return styles.inProgress;
      case 'Completed':
        return styles.completed;
      default:
        return styles.notStarted;
    }
  };

  const getSourceBadgeClass = (): string => {
    switch (task.source) {
      case 'Planner':
        return styles.planner;
      case 'SharePoint':
        return styles.sharepoint;
      default:
        return styles.planner;
    }
  };

  const getPriorityClass = (): string => {
    if (!task.priority) {
      return styles.medium;
    }

    switch (task.priority) {
      case 'Urgent':
        return styles.urgent;
      case 'High':
        return styles.high;
      case 'Medium':
        return styles.medium;
      case 'Low':
        return styles.low;
      default:
        return styles.medium;
    }
  };

  const getPriorityIcon = (): string => {
    if (!task.priority) {
      return '';
    }

    switch (task.priority) {
      case 'Urgent':
        return 'AlertSolid';
      case 'High':
        return 'Important';
      case 'Medium':
        return 'CircleRing';
      case 'Low':
        return 'CircleRing';
      default:
        return '';
    }
  };

  const taskIsOverdue = isOverdue();
  const taskIsDueToday = isDueToday();
  const priorityIcon = getPriorityIcon();
  const cardClass = `${styles.taskCard} ${task.status === 'Completed' ? styles.completed : ''} ${
    taskIsOverdue && highlightOverdue ? styles.overdue : ''
  }`.trim();
  const titleClass = `${styles.taskTitle} ${task.status === 'Completed' ? styles.completed : ''}`.trim();

  return (
    <div className={cardClass}>
      <div className={styles.cardHeader}>
        <h3 className={titleClass}>{escape(task.title)}</h3>
        {showSource && (
          <span className={`${styles.sourceBadge} ${getSourceBadgeClass()}`}>
            {task.source}
          </span>
        )}
      </div>

      <div className={styles.cardBody}>
        <div className={styles.metaInfo}>
          {task.dueDate && (
            <div className={`${styles.metaItem} ${taskIsOverdue ? styles.overdue : ''} ${taskIsDueToday ? styles.dueToday : ''}`}>
              <Icon iconName="Calendar" className={styles.metaIcon} />
              <span>{formatDate(task.dueDate)}</span>
            </div>
          )}

          {task.priority && (
            <div className={`${styles.metaItem} ${styles.priorityBadge} ${getPriorityClass()}`}>
              {priorityIcon && (
                <Icon iconName={priorityIcon} className={styles.metaIcon} />
              )}
              <span>{task.priority}</span>
            </div>
          )}
        </div>

        <div className={`${styles.statusBadge} ${getStatusClass()}`}>
          {task.status === 'Not Started' && <Icon iconName="CircleRing" />}
          {task.status === 'In Progress' && <Icon iconName="ProgressRingDots" />}
          {task.status === 'Completed' && <Icon iconName="CompletedSolid" />}
          <span>{task.status}</span>
        </div>

        {task.planName && (
          <div className={styles.planInfo}>
            <Icon iconName="TaskManager" className={styles.planIcon} />
            <span>{escape(task.planName)}</span>
          </div>
        )}

        {task.listName && (
          <div className={styles.planInfo}>
            <Icon iconName="BulletedList" className={styles.planIcon} />
            <span>{escape(task.listName)}</span>
          </div>
        )}

        {task.assignedTo && task.assignedTo.length > 0 && (
          <div className={styles.metaInfo}>
            <div className={styles.metaItem}>
              <Icon iconName="Contact" className={styles.metaIcon} />
              <span>{task.assignedTo.join(', ')}</span>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default TaskCard;
