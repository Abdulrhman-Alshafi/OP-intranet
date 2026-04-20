import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { Icon } from '@fluentui/react/lib/Icon';
import { Text } from '@fluentui/react/lib/Text';
import { Stack } from '@fluentui/react/lib/Stack';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from '@fluentui/react/lib/DetailsList';

import styles from './UserDevices.module.scss';
import { IHelpDeskDevice, HelpDeskDeviceService } from '../../../services/HelpDeskDeviceService';

export interface IAvailableUserDevicesTabProps {
  devices: IHelpDeskDevice[];
  service: HelpDeskDeviceService;
  requestsListName: string;
  currentUserId: number;
  onDataChange: () => Promise<void>;
  loading: boolean;
}

export const AvailableUserDevicesTab: React.FC<IAvailableUserDevicesTabProps> = (props) => {
  const { devices, service, requestsListName, currentUserId, onDataChange, loading } = props;
  const [requestingDeviceId, setRequestingDeviceId] = React.useState<number | null>(null);
  const [reason, setReason] = React.useState('');
  const [submitting, setSubmitting] = React.useState(false);
  const [errorMsg, setErrorMsg] = React.useState<string | undefined>();
  const [successMsg, setSuccessMsg] = React.useState<string | undefined>();

  const availableDevices = devices.filter(d => d.IsRequestable && d.Status === 'Available');

  const handleSubmitRequest = async (): Promise<void> => {
    if (!requestingDeviceId || !reason) {
      setErrorMsg('Please enter a reason for your request.');
      return;
    }
    setSubmitting(true);
    setErrorMsg(undefined);
    try {
      await service.addRequest(requestsListName, {
        DeviceId: requestingDeviceId,
        RequesterId: currentUserId,
        Title: reason
      });
      setSuccessMsg('Request submitted successfully!');
      setRequestingDeviceId(null);
      setReason('');
      await onDataChange();
    } catch (e) {
      setErrorMsg(e instanceof Error ? e.message : 'Error submitting request');
    } finally {
      setSubmitting(false);
    }
  };

  const closeDialog = (): void => {
    setRequestingDeviceId(null);
    setReason('');
    setErrorMsg(undefined);
  };

  const columns: IColumn[] = [
    { key: 'title', name: 'Device', fieldName: 'Title', minWidth: 160, maxWidth: 280, isResizable: true },
    { key: 'type', name: 'Type', fieldName: 'DeviceType', minWidth: 100, maxWidth: 160, isResizable: true },
    { key: 'sn', name: 'Serial Number', fieldName: 'SerialNumber', minWidth: 120, maxWidth: 200, isResizable: true },
    { key: 'status', name: 'Status', fieldName: 'Status', minWidth: 90, maxWidth: 120 },
    {
      key: 'action', name: '', minWidth: 140, maxWidth: 160,
      onRender: (item: IHelpDeskDevice) => (
        <PrimaryButton
          text="Request Device"
          iconProps={{ iconName: 'Send' }}
          onClick={() => setRequestingDeviceId(item.Id)}
          styles={{ root: { height: 28, padding: '0 10px', fontSize: 12 } }}
        />
      )
    }
  ];

  if (loading && availableDevices.length === 0) {
    return (
      <Stack horizontalAlign="center" styles={{ root: { padding: 32 } }}>
        <Spinner size={SpinnerSize.medium} label="Loading available devices..." />
      </Stack>
    );
  }

  if (availableDevices.length === 0) {
    return (
      <div className={styles.emptyState}>
        <Icon iconName="Devices3" className={styles.emptyStateIcon} />
        <Text variant="large">No devices are currently available for request.</Text>
      </div>
    );
  }

  const selectedDevice = devices.find(d => d.Id === requestingDeviceId);

  return (
    <>
      {successMsg && (
        <MessageBar
          messageBarType={MessageBarType.success}
          onDismiss={() => setSuccessMsg(undefined)}
          styles={{ root: { marginTop: 12 } }}
        >
          {successMsg}
        </MessageBar>
      )}

      <div className={styles.tableScrollWrapper} style={{ marginTop: 16 }}>
        <DetailsList
          items={availableDevices}
          columns={columns}
          setKey="available"
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
        />
      </div>

      <Dialog
        hidden={!requestingDeviceId}
        onDismiss={closeDialog}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Request Device',
          subText: selectedDevice ? `You are requesting: ${selectedDevice.Title}` : ''
        }}
        modalProps={{ isBlocking: true, styles: { main: { maxWidth: 450 } } }}
      >
        <TextField
          label="Reason for request"
          multiline
          rows={3}
          value={reason}
          onChange={(_, v) => setReason(v || '')}
          required
          placeholder="Describe why you need this device"
        />
        {errorMsg && <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginTop: 8 } }}>{errorMsg}</MessageBar>}
        <DialogFooter>
          <PrimaryButton onClick={handleSubmitRequest} text={submitting ? 'Submitting...' : 'Submit'} disabled={submitting} />
          <DefaultButton onClick={closeDialog} text="Cancel" disabled={submitting} />
        </DialogFooter>
      </Dialog>
    </>
  );
};


export interface IAvailableUserDevicesTabProps {
  devices: IHelpDeskDevice[];
  service: HelpDeskDeviceService;
  requestsListName: string;
  currentUserId: number;
  onDataChange: () => Promise<void>;
  loading: boolean;
}

