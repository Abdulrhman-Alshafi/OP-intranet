declare interface IAnnouncementsHubWebPartStrings {
  PropertyPaneDescription: string;
  DisplayGroupName: string;
  ThemeGroupName: string;

  TitleFieldLabel: string;
  LayoutModeLabel: string;
  ItemsToDisplayLabel: string;
  SortOrderLabel: string;
  ShowImagesLabel: string;
  EnableAnimationsLabel: string;
  EnableCategoryColorsLabel: string;
  AccentColorLabel: string;

  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'AnnouncementsHubWebPartStrings' {
  const strings: IAnnouncementsHubWebPartStrings;
  export = strings;
}
