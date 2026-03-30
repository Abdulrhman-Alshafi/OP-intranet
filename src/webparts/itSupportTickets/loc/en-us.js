/* global define */
define([], function () {
  return {
    // Web part title
    WebPartTitle: "IT Support Tickets",
    // Property pane
    PropertyPaneDescription: "Configure the IT Support Tickets web part.",
    GeneralGroupName: "General",
    FunctionBaseUrlFieldLabel: "Azure Function Base URL",
    FunctionBaseUrlFieldDesc: "Base URL of the Azure Function proxy (e.g. https://my-app.azurewebsites.net).",
    FunctionKeyFieldLabel: "Azure Function Key",
    FunctionKeyFieldDesc: "Function key used to authenticate requests to the Azure Function.",
    ItAdminGroupFieldLabel: "IT Admin SharePoint Group Name",
    ItAdminGroupFieldDesc: "Exact name of the SharePoint group whose members can see All Tickets.",
    // Tabs
    TabMyTickets: "My Tickets",
    TabAllTickets: "All Tickets",
    TabNewTicket: "New Ticket",
    // Table columns
    ColumnRef: "Ref",
    ColumnTitle: "Title",
    ColumnStatus: "Status",
    ColumnUrgency: "Urgency",
    ColumnService: "Service",
    ColumnLastUpdate: "Last Update",
    // Status labels
    StatusNew: "New",
    StatusAssigned: "Assigned",
    StatusResolved: "Resolved",
    StatusClosed: "Closed",
    StatusPending: "Pending",
    // Urgency labels
    UrgencyCritical: "Critical",
    UrgencyHigh: "High",
    UrgencyMedium: "Medium",
    UrgencyLow: "Low",
    // Filters
    FilterAllStatuses: "All Statuses",
    // Empty / loading
    NoTicketsMessage: "No tickets found.",
    NoTicketsMyMessage: "You have no support tickets yet.",
    LoadingMessage: "Loading tickets…",
    LoadMoreLabel: "Load more",
    // New ticket form
    FormTitle: "New Support Ticket",
    FieldTitleLabel: "Title",
    FieldTitlePlaceholder: "Short description of the issue",
    FieldDescriptionLabel: "Description",
    FieldDescriptionPlaceholder: "Detailed description of the issue…",
    FieldServiceLabel: "Service",
    FieldServicePlaceholder: "Select a service",
    FieldUrgencyLabel: "Urgency",
    SubmitButtonLabel: "Submit Ticket",
    SubmitSuccessMessage: "Your ticket has been submitted successfully.",
    // Validation
    ValidationTitleRequired: "Title is required.",
    ValidationDescriptionRequired: "Description is required.",
    ValidationServiceRequired: "Please select a service.",
    ValidationUrgencyRequired: "Please select an urgency level.",
    // Errors
    ErrorLoadTickets: "Unable to load tickets. Please try again.",
    ErrorLoadServices: "Unable to load the service list.",
    ErrorSubmitTicket: "Failed to submit ticket. Please try again.",
    ErrorAuthCheck: "Unable to verify your role. Please refresh the page.",
    // Config not set
    ConfigNotSet: "Please configure the web part: set the Azure Function URL and key in the property pane."
  };
});
