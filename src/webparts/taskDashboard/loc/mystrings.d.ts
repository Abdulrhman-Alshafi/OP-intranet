declare interface ITaskDashboardWebPartStrings {
  PropertyPaneDescription: string;
  DataSourcesGroupName: string;
  DisplaySettingsGroupName: string;
  CustomizationGroupName: string;
  PlanSelectionGroupName: string;
  
  TitleFieldLabel: string;
  EnablePlannerLabel: string;
  EnableSharePointLabel: string;
  SharePointListLabel: string;
  SelectPlansLabel: string;
  
  GroupByLabel: string;
  SortByLabel: string;
  ShowCompletedLabel: string;
  HighlightOverdueLabel: string;
  RefreshIntervalLabel: string;
  
  OverdueColorLabel: string;
  InProgressColorLabel: string;
  CompletedColorLabel: string;
  ShowProgressBarsLabel: string;
  ShowSourceBadgesLabel: string;
  
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

declare module 'TaskDashboardWebPartStrings' {
  const strings: ITaskDashboardWebPartStrings;
  export = strings;
}
