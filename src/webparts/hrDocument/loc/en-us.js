/* global define */
define([], function () {
  return {
    // Web part title
    WebPartTitle: "HR Documents",
    // Property pane
    PropertyPaneDescription: "Configure the HR Document web part settings.",
    GeneralGroupName: "General",
    HrLibraryFieldLabel: "Document Library Name",
    HrGroupFieldLabel: "HR SharePoint Group Name",
    // UI labels
    UploadSectionTitle: "Upload HR Document",
    BulkUploadSectionTitle: "Bulk Upload",
    ExcelColumnsHint: "Excel file must contain three columns: FileName, EmployeeEmail, DocumentType (e.g. Offer Letter).",
    NoDocumentsMessage: "No HR documents found.",
    UploadInProgressMessage: "Upload in progress, please wait\u2026",
    // Validation
    ValidationMissingFiles: "Missing files (listed in Excel but not uploaded):",
    ValidationExtraFiles: "Extra files (uploaded but not listed in Excel):",
    ValidationDuplicates: "Duplicate file names detected:",
    ValidationInvalidEmails: "Could not resolve the following employee emails:",
    ValidationFileExists: "A file with this name already exists in the library. Please rename the file or delete the existing one first.",
    ValidationBulkFileExists: "Files already exist in the library (rename or delete them first):",
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
    ColumnDocumentType: "Document Type",
    ColumnCreated: "Created",
    ColumnModified: "Modified",
    // Document type filter
    DocumentTypeLabel: "Type",
    FilterAllTypes: "All Types",
    // Tabs
    TabMyDocuments: "My Documents",
    TabDocuments: "Documents",
    TabUpload: "Upload",
    // Search
    SearchPlaceholderEmployee: "Search by file name\u2026",
    SearchPlaceholderHr: "Search by file name or employee\u2026",
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
