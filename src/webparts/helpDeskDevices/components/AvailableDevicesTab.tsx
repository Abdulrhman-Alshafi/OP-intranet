import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { Stack } from '@fluentui/react/lib/Stack';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { Text } from '@fluentui/react/lib/Text';
import { IHelpDeskDevice, HelpDeskDeviceService } from '../../../services/HelpDeskDeviceService';
import styles from './HelpDeskDevices.module.scss';

export interface IAvailableDevicesTabProps {
  devices: IHelpDeskDevice[];
  service: HelpDeskDeviceService;
  devicesListName: string;
  onDataChange: () => Promise<void>;
  loading: boolean;
}

export const AvailableDevicesTab: React.FC<IAvailableDevicesTabProps> = (props) => {
  const { devices, service, devicesListName, onDataChange, loading } = props;
  const [toggleLoading, setToggleLoading] = React.useState<number | null>(null);
  const [errorMsg, setErrorMsg] = React.useState<string | undefined>();

  const handleToggle = async (deviceId: number, checked: boolean) => {
    setToggleLoading(deviceId);
    setErrorMsg(undefined);
    try {
      await service.updateDeviceIsRequestable(devicesListName, deviceId, checked);
      await onDataChange();
    } catch (e) {
      setErrorMsg(e instanceof Error ? e.message : 'Error updating device');
    } finally {
      setToggleLoading(null);
    }
  };

  const columns: IColumn[] = [
    { key: 'id', name: 'ID', fieldName: 'Id', minWidth: 30, maxWidth: 50 },
    { key: 'title', name: 'Title', fieldName: 'Title', minWidth: 150, maxWidth: 250, isResizable: true },
    { key: 'type', name: 'Device Type', fieldName: 'DeviceType', minWidth: 100, maxWidth: 150, isResizable: true },
    { key: 'sn', name: 'Serial Number', fieldName: 'SerialNumber', minWidth: 100, maxWidth: 200, isResizable: true },
    { key: 'status', name: 'Status', fieldName: 'Status', minWidth: 100, maxWidth: 150 },
    { 
      key: 'isRequestable', name: 'Is Requestable?', minWidth: 120,
      onRender: (item: IHelpDeskDevice) => {
        if (toggleLoading === item.Id) {
          return <Spinner size={1} />;
        }
        return (
          <Toggle 
            checked={item.IsRequestable} 
            onChange={(_, checked) => handleToggle(item.Id, !!checked)} 
            onText="Yes" 
            offText="No" 
            styles={{ root: { margin: 0 } }}
          />
        );
      }
    }
  ];

  if (loading && devices.length === 0) {
    return <Spinner label="Loading devices..." />;
  }

  if (devices.length === 0 && !loading) {
    return (
       <Stack horizontalAlign="center" className={styles.emptyState}>
         <Icon iconName="Devices3" className={styles.emptyStateIcon} />
         <Text variant="large">No devices found in the system.</Text>
       </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { marginTop: 16 } }}>
      {errorMsg && <MessageBar messageBarType={MessageBarType.error}>{errorMsg}</MessageBar>}
      <DetailsList
        items={devices}
        columns={columns}
        setKey="set"
        layoutMode={DetailsListLayoutMode.justified}
        selectionMode={SelectionMode.none}
      />
    </Stack>
  );
};
