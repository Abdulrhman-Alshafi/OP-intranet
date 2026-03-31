import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { Spinner } from '@fluentui/react/lib/Spinner';
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
    { key: 'status', name: 'Status', fieldName: 'Status', minWidth: 100, maxWidth: 150 }
  ];

  if (loading && requests.length === 0) {
    return <Spinner label="Loading your requests..." styles={{ root: { marginTop: 16 } }} />;
  }

  if (requests.length === 0) {
    return (
       <Stack horizontalAlign="center" verticalAlign="center" tokens={{ childrenGap: 16 }} styles={{ root: { minHeight: 200 } }}>
         <Icon iconName="PageList" styles={{ root: { fontSize: 48, color: '#106EBE' } }} />
         <Text variant="large">You have no device requests.</Text>
       </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { marginTop: 16 } }}>
      <DetailsList
        items={requests}
        columns={columns}
        setKey="set"
        layoutMode={DetailsListLayoutMode.justified}
        selectionMode={SelectionMode.none}
      />
    </Stack>
  );
};
