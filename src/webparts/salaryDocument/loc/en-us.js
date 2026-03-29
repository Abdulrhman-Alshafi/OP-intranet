/* global define */
define([], function () {
  return {
    // Web part title
    WebPartTitle: "Salary Documents",
    // Property pane
    PropertyPaneDescription: "Configure the Salary Document web part settings.",
    GeneralGroupName: "General",
    SalaryLibraryFieldLabel: "Document Library Name",
    AccountantGroupFieldLabel: "Accountant SharePoint Group Name",
    // UI labels
    UploadSectionTitle: "Upload Salary Document",
    BulkUploadSectionTitle: "Bulk Upload",
    ExcelColumnsHint: "Excel file must contain two columns: FileName and EmployeeEmail.",
    NoDocumentsMessage: "No salary documents found.",
    UploadInProgressMessage: "Upload in progress, please wait…",
    // Validation
    ValidationMissingFiles: "Missing files (listed in Excel but not uploaded):",
    ValidationExtraFiles: "Extra files (uploaded but not listed in Excel):",
    ValidationDuplicates: "Duplicate file names detected:",
    ValidationInvalidEmails: "Could not resolve the following employee emails:",
    // Actions
    RetryFailedLabel: "Retry failed uploads",
    DownloadErrorReportLabel: "Download error report",
    DeleteConfirmTitle: "Delete Document",
    DeleteConfirmMessage: "Are you sure you want to permanently delete '{name}'? This cannot be undone.",
    DeleteButtonLabel: "Delete",
    CancelButtonLabel: "Cancel",
    // Column headers
    ColumnFileName: "File Name",
    ColumnEmployee: "Employee",
    ColumnCreated: "Created",
    ColumnModified: "Modified",
    // Tabs
    TabMySalary: "My Salary",
    TabDocuments: "Documents",
    TabUpload: "Upload",
    // Search
    SearchPlaceholderEmployee: "Search by file name…",
    SearchPlaceholderAccountant: "Search by file name or employee…",
    // Upload panel
    SelectFileLabel: "Select file",
    SelectEmployeeLabel: "Select employee",
    UploadButtonLabel: "Upload",
    SelectExcelLabel: "Select Excel manifest (.xlsx)",
    SelectDocumentsLabel: "Select documents",
    ProceedBulkLabel: "Proceed with bulk upload",
    LoadMoreLabel: "Load more"
  };
});
