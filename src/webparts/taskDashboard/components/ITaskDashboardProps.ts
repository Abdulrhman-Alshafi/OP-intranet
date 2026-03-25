import { MSGraphClientV3 } from '@microsoft/sp-http';
import { SPHttpClient } from '@microsoft/sp-http';
import { ITask } from '../../../services/TaskService';

/**
 * Selected plan structure for PropertyFieldCollectionData
 */
export interface ISelectedPlan {
  id: string;
  title: string;
}

/**
 * Props passed from WebPart to TaskDashboard component
 */
export interface ITaskDashboardProps {
  // WebPart properties
  title: string;
  enablePlanner: boolean;
  enableSharePoint: boolean;
  sharePointListId: string;
  selectedPlans: ISelectedPlan[];
  groupBy: 'status' | 'source' | 'dueDate' | 'none';
  sortBy: 'dueDate' | 'priority' | 'title';
  showCompleted: boolean;
  highlightOverdue: boolean;
  refreshInterval: number;
  overdueColor: string;
  inProgressColor: string;
  completedColor: string;
  showProgressBars: boolean;
  showSourceBadges: boolean;

  // Context services
  graphClient: MSGraphClientV3;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  
  // Theme and environment
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}

/**
 * Filter state for task filtering
 */
export interface ITaskFilter {
  sources: Array<'Planner' | 'SharePoint'>;
  statuses: Array<'Not Started' | 'In Progress' | 'Completed'>;
}

/**
 * Props for TaskCard component
 */
export interface ITaskCardProps {
  task: ITask;
  showProgress: boolean;
  showSource: boolean;
  overdueColor: string;
  inProgressColor: string;
  completedColor: string;
  highlightOverdue: boolean;
}
