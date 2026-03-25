export interface IServiceItem {
  imageUrl?: string;
  title: string;
  description?: string;
  linkText?: string;
  linkUrl?: string;
  cardColor?: string;
}

export interface IServicesGridProps {
  sectionTitle: string;
  services: IServiceItem[];
  columns: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
