declare interface ISalaryDocumentWebPartStrings {
  // Web part title
  WebPartTitle: string;
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
  DeleteConfirmTitle: string;
  DeleteConfirmMessage: string;
  DeleteButtonLabel: string;
  CancelButtonLabel: string;
  // Column headers
  ColumnFileName: string;
  ColumnEmployee: string;
  ColumnPayPeriod: string;
  ColumnCreated: string;
  ColumnModified: string;
  // Period filter
  PayPeriodYearLabel: string;
  PayPeriodMonthLabel: string;
  FilterAllYears: string;
  FilterAllMonths: string;
  // Tabs
  TabMySalary: string;
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
