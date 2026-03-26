declare interface IAnnouncementsWebPartStrings {
  PropertyPaneDescription: string;
  GeneralGroupName: string;
  DataSourcesGroupName: string;
  DisplaySettingsGroupName: string;

  TitleFieldLabel: string;
  AnnouncementsListLabel: string;
  EnableCreateAnnouncementLabel: string;
  EnableReactionsLabel: string;
  EnableCommentsLabel: string;
  ItemsPerPageLabel: string;

  AppLocalEnvironmentSharePoint: string;
  AppSharePointEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'AnnouncementsWebPartStrings' {
  const strings: IAnnouncementsWebPartStrings;
  export = strings;
}
