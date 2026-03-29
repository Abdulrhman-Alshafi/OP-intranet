import * as React from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  ConstrainMode
} from '@fluentui/react/lib/DetailsList';
import { Icon } from '@fluentui/react/lib/Icon';
import { Link } from '@fluentui/react/lib/Link';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { DefaultButton, PrimaryButton, IconButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';

import * as strings from 'SalaryDocumentWebPartStrings';
import styles from './SalaryDocument.module.scss';
import { ISalaryDocument } from '../../../services/SalaryDocumentService';

// ── File icon map ─────────────────────────────────────────────────────────────

const EXTENSION_ICONS: Record<string, string> = {
  pdf: 'PDF',
  doc: 'WordDocument',
  docx: 'WordDocument',
  xls: 'ExcelDocument',
  xlsx: 'ExcelDocument',
  ppt: 'PowerPointDocument',
  pptx: 'PowerPointDocument',
  png: 'FileImage',
  jpg: 'FileImage',
  jpeg: 'FileImage',
  zip: 'ZipFolder'
};

function getFileIcon(extension: string): string {
  return EXTENSION_ICONS[extension.toLowerCase()] ?? 'Document';
}

function formatDate(iso: string): string {
  if (!iso) return '';
  try {
    return new Date(iso).toLocaleDateString(undefined, {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    });
  } catch {
    return iso;
  }
}

const MONTH_ABBREVS = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function formatPeriod(year: number, month: number): string {
  if (!year && !month) return '';
  if (year && !month) return String(year);
  if (!year && month) return MONTH_ABBREVS[month] || '';
  return `${MONTH_ABBREVS[month] || ''} ${year}`;
}

// ── Props ─────────────────────────────────────────────────────────────────────

export interface ISalaryTableProps {
  items: ISalaryDocument[];
  showEmployeeColumn: boolean;
  loading: boolean;
  nextLink: string | undefined;
  onLoadMore: () => void;
  canDelete?: boolean;
  onDeleteRequest?: (item: ISalaryDocument) => Promise<void>;
}

// ── Component ─────────────────────────────────────────────────────────────────

export const SalaryTable: React.FC<ISalaryTableProps> = ({
  items,
  showEmployeeColumn,
  loading,
  nextLink,
  onLoadMore,
  canDelete,
  onDeleteRequest
}) => {
  const [pendingDeleteItem, setPendingDeleteItem] = React.useState<ISalaryDocument | null>(null);
  const [isDeleting, setIsDeleting] = React.useState<boolean>(false);

  const handleConfirmDelete = React.useCallback(async () => {
    if (!pendingDeleteItem || !onDeleteRequest) return;
    setIsDeleting(true);
    try {
      await onDeleteRequest(pendingDeleteItem);
    } finally {
      setIsDeleting(false);
      setPendingDeleteItem(null);
    }
  }, [pendingDeleteItem, onDeleteRequest]);

  const columns = React.useMemo<IColumn[]>(() => {
    const cols: IColumn[] = [
      {
        key: 'icon',
        name: '',
        minWidth: 24,
        maxWidth: 24,
        className: styles.fileIconCell,
        onRender: (item: ISalaryDocument) => (
          <Icon iconName={getFileIcon(item.extension)} />
        )
      },
      {
        key: 'name',
        name: strings.ColumnFileName,
        fieldName: 'name',
        minWidth: 120,
        maxWidth: 300,
        isResizable: true,
        isMultiline: false,
        onRender: (item: ISalaryDocument) => (
          <Link
            href={item.serverRelativeUrl}
            target="_blank"
            rel="noreferrer noopener"
            title={item.name}
            styles={{ root: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', display: 'block', maxWidth: '100%' } }}
          >
            {item.name}
          </Link>
        )
      }
    ];

    if (showEmployeeColumn) {
      cols.push({
        key: 'employee',
        name: strings.ColumnEmployee,
        fieldName: 'employeeDisplayName',
        minWidth: 150,
        maxWidth: 250,
        isResizable: true,
        onRender: (item: ISalaryDocument) => (
          <Text variant="small" title={item.employeeDisplayName || item.employeeEmail}
            styles={{ root: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', display: 'block' } }}>
            {item.employeeDisplayName || item.employeeEmail}
          </Text>
        )
      });
    }

    cols.push({
      key: 'payPeriod',
      name: strings.ColumnPayPeriod,
      minWidth: 90,
      maxWidth: 110,
      isResizable: true,
      onRender: (item: ISalaryDocument) => (
        <Text variant="small">{formatPeriod(item.payPeriodYear, item.payPeriodMonth)}</Text>
      )
    });

    cols.push(
      {
        key: 'created',
        name: strings.ColumnCreated,
        fieldName: 'created',
        minWidth: 100,
        maxWidth: 130,
        isResizable: true,
        onRender: (item: ISalaryDocument) => (
          <Text variant="small">{formatDate(item.created)}</Text>
        )
      },
      {
        key: 'modified',
        name: strings.ColumnModified,
        fieldName: 'modified',
        minWidth: 100,
        maxWidth: 130,
        isResizable: true,
        onRender: (item: ISalaryDocument) => (
          <Text variant="small">{formatDate(item.modified)}</Text>
        )
      }
    );

    if (canDelete) {
      cols.push({
        key: 'delete',
        name: '',
        minWidth: 36,
        maxWidth: 36,
        onRender: (item: ISalaryDocument) => (
          <IconButton
            iconProps={{ iconName: 'Delete' }}
            title="Delete document"
            ariaLabel="Delete document"
            onClick={() => setPendingDeleteItem(item)}
            styles={{
              root: { color: '#a4262c' },
              rootHovered: { color: '#750b1c', backgroundColor: 'rgba(164,38,44,0.08)' }
            }}
          />
        )
      });
    }

    return cols;
  }, [showEmployeeColumn, canDelete]);

  if (!loading && items.length === 0) {
    return (
      <Stack horizontalAlign="center" className={styles.emptyState}>
        <Icon iconName="DocumentSearch" className={styles.emptyStateIcon} />
        <Text variant="large">{strings.NoDocumentsMessage}</Text>
      </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <div className={styles.tableScrollWrapper}>
        <DetailsList
          items={items}
          columns={columns}
          layoutMode={DetailsListLayoutMode.fixedColumns}
          constrainMode={ConstrainMode.unconstrained}
          selectionMode={SelectionMode.none}
          isHeaderVisible={true}
          compact={false}
        />
      </div>
      {loading && (
        <Stack horizontalAlign="center" styles={{ root: { padding: '8px 0' } }}>
          <Spinner size={SpinnerSize.small} />
        </Stack>
      )}
      {!loading && nextLink && (
        <Stack horizontalAlign="center" className={styles.loadMoreWrapper}>
          <DefaultButton text={strings.LoadMoreLabel} onClick={onLoadMore} />
        </Stack>
      )}

      <Dialog
        hidden={pendingDeleteItem === null}
        onDismiss={() => { if (!isDeleting) setPendingDeleteItem(null); }}
        dialogContentProps={{
          type: DialogType.normal,
          title: strings.DeleteConfirmTitle,
          subText: pendingDeleteItem
            ? strings.DeleteConfirmMessage.replace('{name}', pendingDeleteItem.name)
            : ''
        }}
        modalProps={{ isBlocking: isDeleting }}
      >
        <DialogFooter>
          <PrimaryButton
            text={isDeleting ? '...' : strings.DeleteButtonLabel}
            onClick={handleConfirmDelete}
            disabled={isDeleting}
            styles={{ root: { backgroundColor: '#a4262c', borderColor: '#a4262c' } }}
          />
          <DefaultButton
            text={strings.CancelButtonLabel}
            onClick={() => setPendingDeleteItem(null)}
            disabled={isDeleting}
          />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
};

export default SalaryTable;
