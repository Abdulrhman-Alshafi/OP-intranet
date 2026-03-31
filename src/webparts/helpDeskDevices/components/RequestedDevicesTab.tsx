import * as React from 'react';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from '@fluentui/react/lib/DetailsList';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { Text } from '@fluentui/react/lib/Text';
import { IHelpDeskDevice, IHelpDeskDeviceRequest, HelpDeskDeviceService } from '../../../services/HelpDeskDeviceService';
import styles from './HelpDeskDevices.module.scss';

export interface IRequestedDevicesTabProps {
  requests: IHelpDeskDeviceRequest[];
  devices: IHelpDeskDevice[];
  service: HelpDeskDeviceService;
  requestsListName: string;
  onDataChange: () => Promise<void>;
  loading: boolean;
}

export const RequestedDevicesTab: React.FC<IRequestedDevicesTabProps> = (props) => {
  const { requests, devices, service, requestsListName, onDataChange, loading } = props;
  const [actingOn, setActingOn] = React.useState<number | null>(null);
  const [errorMsg, setErrorMsg] = React.useState<string | undefined>();

  const handleAction = async (requestId: number, newStatus: string) => {
    setActingOn(requestId);
    setErrorMsg(undefined);
    try {
      await service.updateRequestStatus(requestsListName, requestId, newStatus);
      await onDataChange();
    } catch (e) {
      setErrorMsg(e instanceof Error ? e.message : 'Error updating request');
    } finally {
      setActingOn(null);
    }
  };

  const columns: IColumn[] = [
    { key: 'id', name: 'ID', fieldName: 'Id', minWidth: 30, maxWidth: 50 },
    { key: 'reason', name: 'Reason', fieldName: 'Title', minWidth: 150, maxWidth: 300, isResizable: true },
    { key: 'requester', name: 'Requester', fieldName: 'RequesterName', minWidth: 100, maxWidth: 200, isResizable: true },
    { 
      key: 'device', name: 'Device', minWidth: 100, maxWidth: 200, isResizable: true,
      onRender: (item: IHelpDeskDeviceRequest) => {
        const d = devices.find(x => x.Id === item.DeviceId);
        return d ? `${d.Title} (${d.SerialNumber || 'No SN'})` : `Device ID: ${item.DeviceId}`;
      }
    },
    { key: 'status', name: 'Status', fieldName: 'Status', minWidth: 100, maxWidth: 150 },
    { 
      key: 'actions', name: 'Actions', minWidth: 200,
      onRender: (item: IHelpDeskDeviceRequest) => {
        if (item.Status !== 'Pending') {
          return <span>-</span>;
        }
        if (actingOn === item.Id) {
          return <Spinner size={1} />;
        }
        return (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton text="Approve" onClick={() => handleAction(item.Id, 'Approved')} />
            <DefaultButton text="Reject" onClick={() => handleAction(item.Id, 'Rejected')} />
          </Stack>
        );
      }
    }
  ];

  if (loading && requests.length === 0) {
    return <Spinner label="Loading requests..." />;
  }

  if (requests.length === 0 && !loading) {
    return (
       <Stack horizontalAlign="center" verticalAlign="center" tokens={{ childrenGap: 16 }} styles={{ root: { minHeight: 200 } }}>
         <Icon iconName="InboxCheck" styles={{ root: { fontSize: 48, color: '#106EBE' } }} />
         <Text variant="large">No pending device requests.</Text>
       </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { marginTop: 16 } }}>
      {errorMsg && <MessageBar messageBarType={MessageBarType.error}>{errorMsg}</MessageBar>}
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
