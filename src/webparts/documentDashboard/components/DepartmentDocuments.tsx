import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { DetailsList, IColumn, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Link } from '@fluentui/react/lib/Link';
import { Icon } from '@fluentui/react/lib/Icon';
import { ShimmeredDetailsList } from '@fluentui/react/lib/ShimmeredDetailsList';
import { ISharePointDocument } from '../../../services/sharepointService';
import { IDepartmentDocumentsProps } from './IDashboardProps';

const getFileIconName = (extension: string): string => {
  switch (extension) {
    case 'pdf':
      return 'PDF';
    case 'doc':
    case 'docx':
      return 'WordDocument';
    case 'xls':
    case 'xlsx':
      return 'ExcelDocument';
    case 'ppt':
    case 'pptx':
      return 'PowerPointDocument';
    case 'txt':
      return 'TextDocument';
    default:
      return 'Page';
  }
};

const formatDate = (dateValue: string): string => {
  const date = new Date(dateValue);
  return date.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
};

const DepartmentDocuments: React.FC<IDepartmentDocumentsProps> = ({ departmentFiles, loading, error, siteUrl }) => {
  const columns: IColumn[] = React.useMemo(
    () => [
      {
        key: 'icon',
        name: '',
        minWidth: 28,
        maxWidth: 28,
        isResizable: false,
        onRender: (item: ISharePointDocument) => (
          <Icon iconName={getFileIconName(item.extension)} aria-label={item.extension || 'file'} />
        )
      },
      {
        key: 'name',
        name: 'File name',
        fieldName: 'name',
        minWidth: 240,
        maxWidth: 520,
        isResizable: true,
        onRender: (item: ISharePointDocument) => (
          <Link href={`${siteUrl}${item.serverRelativeUrl}`} target="_blank" rel="noreferrer">
            {item.name}
          </Link>
        )
      },
      {
        key: 'created',
        name: 'Created',
        fieldName: 'created',
        minWidth: 120,
        maxWidth: 180,
        isResizable: true,
        onRender: (item: ISharePointDocument) => formatDate(item.created)
      }
    ],
    [siteUrl]
  );

  if (error) {
    return (
      <MessageBar messageBarType={MessageBarType.error} isMultiline>
        {error}
      </MessageBar>
    );
  }

  if (!loading && departmentFiles.length === 0) {
    return (
      <MessageBar messageBarType={MessageBarType.info} isMultiline={false}>
        No department documents are currently available.
      </MessageBar>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      {loading && (
        <ShimmeredDetailsList
          items={[]}
          columns={columns}
          selectionMode={SelectionMode.none}
          enableShimmer={true}
          ariaLabelForShimmer="Loading department documents"
        />
      )}

      {!loading && departmentFiles.map((department) => (
        <Stack key={department.departmentName} tokens={{ childrenGap: 8 }}>
          <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
            {department.departmentName}
          </Text>

          {department.files.length > 0 ? (
            <DetailsList
              items={department.files}
              columns={columns}
              selectionMode={SelectionMode.none}
            />
          ) : (
            <Text variant="small">No files available in this department.</Text>
          )}
        </Stack>
      ))}
    </Stack>
  );
};

export default DepartmentDocuments;
