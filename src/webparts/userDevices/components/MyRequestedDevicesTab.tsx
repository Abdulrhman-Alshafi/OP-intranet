import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Stack } from '@fluentui/react/lib/Stack';
import { Icon } from '@fluentui/react/lib/Icon';
import { Text } from '@fluentui/react/lib/Text';
import { IHelpDeskDevice, IHelpDeskDeviceRequest } from '../../../services/HelpDeskDeviceService';
import styles from './UserDevices.module.scss';

export interface IMyRequestedDevicesTabProps {
  requests: IHelpDeskDeviceRequest[];
  devices: IHelpDeskDevice[];
  loading: boolean;
}

export const MyRequestedDevicesTab: React.FC<IMyRequestedDevicesTabProps> = (props) => {
  const { requests, devices, loading } = props;

  const columns: IColumn[] = [
    { key: 'id', name: 'Request ID', fieldName: 'Id', minWidth: 50, maxWidth: 80 },
    {
      key: 'device', name: 'Device', minWidth: 150, maxWidth: 250, isResizable: true,
      onRender: (item: IHelpDeskDeviceRequest) => {
        const d = devices.find(x => x.Id === item.DeviceId);
        return d ? d.Title : `Device ID: ${item.DeviceId}`;
      }
    },
    { key: 'reason', name: 'Reason', fieldName: 'Title', minWidth: 150, maxWidth: 300, isResizable: true },
    {
      key: 'status', name: 'Status', fieldName: 'Status', minWidth: 100, maxWidth: 150,
      onRender: (item: IHelpDeskDeviceRequest) => {
        const color =
          item.Status === 'Approved' ? '#107c10' :
          item.Status === 'Rejected' ? '#a4262c' : '#605e5c';
        return <span style={{ color, fontWeight: 600 }}>{item.Status}</span>;
      }
    }
  ];

  if (loading && requests.length === 0) {
    return (
      <Stack horizontalAlign="center" styles={{ root: { padding: 32 } }}>
        <Spinner size={SpinnerSize.medium} label="Loading your requests..." />
      </Stack>
    );
  }

  if (requests.length === 0) {
    return (
      <div className={styles.emptyState}>
        <Icon iconName="PageList" className={styles.emptyStateIcon} />
        <Text variant="large">You have no device requests yet.</Text>
      </div>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { marginTop: 16 } }}>
      <div className={styles.tableScrollWrapper}>
        <DetailsList
          items={requests}
          columns={columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
        />
      </div>
    </Stack>
  );
};


export interface IMyRequestedDevicesTabProps {
  requests: IHelpDeskDeviceRequest[];
  devices: IHelpDeskDevice[];
  loading: boolean;
}

