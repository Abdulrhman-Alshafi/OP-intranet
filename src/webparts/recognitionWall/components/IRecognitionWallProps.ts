import { SPHttpClient } from '@microsoft/sp-http';
import { IKudosPost, ITopEmployee, IComment, IEmployeeOfMonth } from '../../../services/RecognitionService';

/**
 * Props passed from WebPart to RecognitionWall component
 */
export interface IRecognitionWallProps {
  title: string;
  refreshInterval: number;
  showEmployeeOfMonth: boolean;
  postsPerPage: number;
  spHttpClient: SPHttpClient;
  siteUrl: string;
  currentUserId: number;
  currentUserEmail: string;
  currentUserName: string;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  isAdmin: boolean;
}

/**
 * Props for the KudosPost card component
 */
export interface IKudosPostProps {
  post: IKudosPost;
  currentUserId: number;
  hasLiked: boolean;
  isAdmin: boolean;
  onToggleLike: (postId: number) => Promise<void>;
  onAddComment: (postId: number, text: string) => void;
  onLoadComments: (postId: number) => Promise<IComment[]>;
  onDelete: (postId: number) => Promise<void>;
}

/**
 * Props for the CommentSection component
 */
export interface ICommentSectionProps {
  postId: number;
  comments: IComment[];
  loading: boolean;
  onAddComment: (postId: number, text: string) => void;
}

/**
 * Props for the KudosForm (create new kudos) component
 */
export interface IKudosFormProps {
  isOpen: boolean;
  onDismiss: () => void;
  onSubmit: (recipientEmail: string, recipientName: string, message: string, category: string) => void;
  siteUrl: string;
  spHttpClient: SPHttpClient;
}

/**
 * Props for the EmployeeHighlight component
 */
export interface IEmployeeHighlightProps {
  employee: ITopEmployee | IEmployeeOfMonth;
  siteUrl: string;
  isAdmin?: boolean;
  onRemove?: () => void;
}
