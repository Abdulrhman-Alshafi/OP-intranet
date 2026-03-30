declare interface IPollsSurveyWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;

  TitleFieldLabel: string;
  RefreshIntervalLabel: string;
  AllowAnonymousLabel: string;
  PollsPerPageLabel: string;

  AutoHideDaysLabel: string;

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

declare module 'PollsSurveyWebPartStrings' {
  const strings: IPollsSurveyWebPartStrings;
  export = strings;
}
