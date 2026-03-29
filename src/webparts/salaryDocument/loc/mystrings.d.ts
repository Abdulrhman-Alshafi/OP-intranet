declare interface ISalaryDocumentWebPartStrings {
  // Property pane
  PropertyPaneDescription: string;
  GeneralGroupName: string;
  SalaryLibraryFieldLabel: string;
  AccountantGroupFieldLabel: string;
  // UI labels
  UploadSectionTitle: string;
  BulkUploadSectionTitle: string;
  ExcelColumnsHint: string;
  NoDocumentsMessage: string;
  UploadInProgressMessage: string;
  // Validation
  ValidationMissingFiles: string;
  ValidationExtraFiles: string;
  ValidationDuplicates: string;
  ValidationInvalidEmails: string;
  // Actions
  RetryFailedLabel: string;
  DownloadErrorReportLabel: string;
  // Column headers
  ColumnFileName: string;
  ColumnEmployee: string;
  ColumnCreated: string;
  ColumnModified: string;
  // Tabs
  TabDocuments: string;
  TabUpload: string;
  // Search
  SearchPlaceholderEmployee: string;
  SearchPlaceholderAccountant: string;
  // Upload panel
  SelectFileLabel: string;
  SelectEmployeeLabel: string;
  UploadButtonLabel: string;
  SelectExcelLabel: string;
  SelectDocumentsLabel: string;
  ProceedBulkLabel: string;
  LoadMoreLabel: string;
}

declare module 'SalaryDocumentWebPartStrings' {
  const strings: ISalaryDocumentWebPartStrings;
  export = strings;
}
