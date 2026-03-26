import * as React from 'react';
import { IColumn, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { Icon } from '@fluentui/react/lib/Icon';
import { Link } from '@fluentui/react/lib/Link';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { ShimmeredDetailsList } from '@fluentui/react/lib/ShimmeredDetailsList';
import { Text } from '@fluentui/react/lib/Text';
import { ISharePointDocument } from '../../../services/sharepointService';
import { IPersonalDocumentsProps } from './IDashboardProps';

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
    case 'zip':
      return 'ZipFolder';
    default:
      return 'Page';
  }
};

const formatDate = (dateValue: string): string => {
  const date = new Date(dateValue);
  return date.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
};

const PersonalDocuments: React.FC<IPersonalDocumentsProps> = ({ files, loading, error, siteUrl }) => {
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

  if (!loading && files.length === 0) {
    return (
      <MessageBar messageBarType={MessageBarType.info} isMultiline={false}>
        No personal documents are available for your account.
      </MessageBar>
    );
  }

  return (
    <>
      <ShimmeredDetailsList
        items={files}
        columns={columns}
        selectionMode={SelectionMode.none}
        enableShimmer={loading}
        ariaLabelForShimmer="Loading personal documents"
      />
      {!loading && files.length > 0 && (
        <Text variant="small" styles={{ root: { marginTop: 8 } }}>
          {files.length} document{files.length === 1 ? '' : 's'}
        </Text>
      )}
    </>
  );
};

export default PersonalDocuments;
