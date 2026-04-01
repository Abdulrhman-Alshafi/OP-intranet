export interface IHelpDeskButton {
  label: string;
  iconName: string;
  url: string;
  openInNewTab: boolean;
}

export interface IHelpDeskTile {
  title: string;
  linkText: string;
  linkUrl: string;
  imageUrl: string;
  openInNewTab: boolean;
}

export interface IItHelpDeskHeroProps {
  title: string;
  titleColor: string;
  description: string;
  backgroundColor: string;
  leftColumnWidthPercent: string;
  buttons: IHelpDeskButton[];
  tiles: IHelpDeskTile[];
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
