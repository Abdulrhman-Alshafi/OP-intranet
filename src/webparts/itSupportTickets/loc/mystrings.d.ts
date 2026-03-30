declare interface IItSupportTicketsWebPartStrings {
  WebPartTitle: string;
  PropertyPaneDescription: string;
  GeneralGroupName: string;
  FunctionBaseUrlFieldLabel: string;
  FunctionBaseUrlFieldDesc: string;
  FunctionKeyFieldLabel: string;
  FunctionKeyFieldDesc: string;
  ItAdminGroupFieldLabel: string;
  ItAdminGroupFieldDesc: string;
  TabMyTickets: string;
  TabAllTickets: string;
  TabNewTicket: string;
  ColumnRef: string;
  ColumnTitle: string;
  ColumnStatus: string;
  ColumnUrgency: string;
  ColumnService: string;
  ColumnLastUpdate: string;
  StatusNew: string;
  StatusAssigned: string;
  StatusResolved: string;
  StatusClosed: string;
  StatusPending: string;
  UrgencyCritical: string;
  UrgencyHigh: string;
  UrgencyMedium: string;
  UrgencyLow: string;
  FilterAllStatuses: string;
  NoTicketsMessage: string;
  NoTicketsMyMessage: string;
  LoadingMessage: string;
  LoadMoreLabel: string;
  FormTitle: string;
  FieldTitleLabel: string;
  FieldTitlePlaceholder: string;
  FieldDescriptionLabel: string;
  FieldDescriptionPlaceholder: string;
  FieldServiceLabel: string;
  FieldServicePlaceholder: string;
  FieldUrgencyLabel: string;
  SubmitButtonLabel: string;
  SubmitSuccessMessage: string;
  ValidationTitleRequired: string;
  ValidationDescriptionRequired: string;
  ValidationServiceRequired: string;
  ValidationUrgencyRequired: string;
  ErrorLoadTickets: string;
  ErrorLoadServices: string;
  ErrorSubmitTicket: string;
  ErrorAuthCheck: string;
  ConfigNotSet: string;
}

declare module 'ItSupportTicketsWebPartStrings' {
  const strings: IItSupportTicketsWebPartStrings;
  export = strings;
}
