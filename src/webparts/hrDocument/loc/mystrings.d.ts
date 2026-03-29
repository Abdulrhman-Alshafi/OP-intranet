declare interface IHrDocumentWebPartStrings {
  // Web part title
  WebPartTitle: string;
  // Property pane
  PropertyPaneDescription: string;
  GeneralGroupName: string;
  HrLibraryFieldLabel: string;
  HrGroupFieldLabel: string;
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
  ValidationFileExists: string;
  ValidationBulkFileExists: string;
  // Actions
  RetryFailedLabel: string;
  DownloadErrorReportLabel: string;
  DeleteConfirmTitle: string;
  DeleteConfirmMessage: string;
  DeleteButtonLabel: string;
  CancelButtonLabel: string;
  // Column headers
  ColumnFileName: string;
  ColumnEmployee: string;
  ColumnDocumentType: string;
  ColumnCreated: string;
  ColumnModified: string;
  // Document type filter
  DocumentTypeLabel: string;
  FilterAllTypes: string;
  // Tabs
  TabMyDocuments: string;
  TabDocuments: string;
  TabUpload: string;
  // Search
  SearchPlaceholderEmployee: string;
  SearchPlaceholderHr: string;
  // Upload panel
  SelectFileLabel: string;
  SelectEmployeeLabel: string;
  UploadButtonLabel: string;
  SelectExcelLabel: string;
  SelectDocumentsLabel: string;
  ProceedBulkLabel: string;
  LoadMoreLabel: string;
}

declare module 'HrDocumentWebPartStrings' {
  const strings: IHrDocumentWebPartStrings;
  export = strings;
}
