import { SPHttpClient } from '@microsoft/sp-http';

export interface IPollsSurveyProps {
  /** Web part title */
  title: string;
  /** Auto-refresh interval in seconds */
  refreshInterval: number;
  /** Whether anonymous voting is allowed */
  allowAnonymous: boolean;
  /** Number of polls to display per page */
  pollsPerPage: number;
  /** SharePoint HTTP client instance */
  spHttpClient: SPHttpClient;
  /** Absolute URL of the current site */
  siteUrl: string;
  /** Current user's SharePoint numeric ID (for Person field) */
  currentUserId: number;
  /** Current user's display name */
  currentUserDisplayName: string;
  /** Current user's email address */
  currentUserEmail: string;
  /** Whether the current theme is dark */
  isDarkTheme: boolean;
  /** Whether running inside Microsoft Teams */
  hasTeamsContext: boolean;
}
