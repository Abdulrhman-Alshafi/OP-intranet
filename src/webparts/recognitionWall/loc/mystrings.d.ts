declare interface IRecognitionWallWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;

  TitleFieldLabel: string;
  RefreshIntervalLabel: string;
  ShowEmployeeOfMonthLabel: string;
  PostsPerPageLabel: string;

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

declare module 'RecognitionWallWebPartStrings' {
  const strings: IRecognitionWallWebPartStrings;
  export = strings;
}
