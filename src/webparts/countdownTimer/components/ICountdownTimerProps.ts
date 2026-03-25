export interface ICountdownTimerProps {
  eventTitle: string;
  eventDescription: string;
  targetDate: string;
  backgroundImage: string;
  overlayOpacity: number;
  layout: 'horizontal' | 'compact' | 'banner';
  showDays: boolean;
  showHours: boolean;
  showMinutes: boolean;
  showSeconds: boolean;
  completedMessage: string;
  accentColor: string;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
}
