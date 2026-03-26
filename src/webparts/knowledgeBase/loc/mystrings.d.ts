declare interface IKnowledgeBaseWebPartStrings {
  PropertyPaneDescription: string;
  GeneralGroupName: string;
  DataGroupName: string;
  DisplayGroupName: string;

  TitleFieldLabel: string;
  ListNameFieldLabel: string;
  DefaultCategoryFieldLabel: string;
  ItemsPerPageFieldLabel: string;
  EnableSearchFieldLabel: string;
  EnableCategoryFilterFieldLabel: string;

  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'KnowledgeBaseWebPartStrings' {
  const strings: IKnowledgeBaseWebPartStrings;
  export = strings;
}
