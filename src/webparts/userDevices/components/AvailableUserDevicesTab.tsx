import * as React from 'react';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { Text } from '@fluentui/react/lib/Text';
import { Stack } from '@fluentui/react/lib/Stack';

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

  // Filter to only show devices that are Available AND isRequestable
  const availableDevices = devices.filter(d => d.IsRequestable && d.Status === 'Available');

  const handleSubmitRequest = async () => {
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

  const closeDialog = () => {
    setRequestingDeviceId(null);
    setReason('');
    setErrorMsg(undefined);
  };

  if (loading && devices.length === 0) {
    return <Spinner label="Loading available devices..." styles={{ root: { marginTop: 16 } }} />;
  }

  if (availableDevices.length === 0) {
    return (
      <Stack horizontalAlign="center" verticalAlign="center" tokens={{ childrenGap: 16 }} styles={{ root: { minHeight: 200 } }}>
        <Icon iconName="Devices3" styles={{ root: { fontSize: 48, color: '#106EBE' } }} />
        <Text variant="large">No devices are currently available for request.</Text>
      </Stack>
    );
  }

  const selectedDevice = devices.find(d => d.Id === requestingDeviceId);

  return (
    <>
      {successMsg && <MessageBar messageBarType={MessageBarType.success} styles={{ root: { marginTop: 16 } }}>{successMsg}</MessageBar>}
      <div className={styles.grid}>
        {availableDevices.map(device => (
          <div key={device.Id} className={styles.card}>
            <div>
              <p className={styles.cardTitle}>{device.Title}</p>
              <p className={styles.cardSubText}>{device.DeviceType}</p>
              {device.SerialNumber && <p className={styles.cardSubText}>SN: {device.SerialNumber}</p>}
            </div>
            <PrimaryButton 
              text="Request Device" 
              onClick={() => setRequestingDeviceId(device.Id)} 
            />
          </div>
        ))}
      </div>

      <Dialog
        hidden={!requestingDeviceId}
        onDismiss={closeDialog}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Request Device'
        }}
        modalProps={{
          isBlocking: true,
          styles: { main: { maxWidth: 450 } }
        }}
      >
        <p>You are requesting: <strong>{selectedDevice?.Title}</strong></p>
        <TextField
          label="Reason for request"
          multiline
          rows={3}
          value={reason}
          onChange={(_, v) => setReason(v || '')}
          required
        />
        {errorMsg && <MessageBar messageBarType={MessageBarType.error}>{errorMsg}</MessageBar>}
        <DialogFooter>
          <PrimaryButton onClick={handleSubmitRequest} text="Submit" disabled={submitting} />
          <DefaultButton onClick={closeDialog} text="Cancel" disabled={submitting} />
        </DialogFooter>
      </Dialog>
    </>
  );
};
