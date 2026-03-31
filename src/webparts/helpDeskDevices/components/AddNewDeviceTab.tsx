import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { TextField } from '@fluentui/react/lib/TextField';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { HelpDeskDeviceService } from '../../../services/HelpDeskDeviceService';

export interface IAddNewDeviceTabProps {
  service: HelpDeskDeviceService;
  devicesListName: string;
  onDeviceAdded: () => Promise<void>;
}

export const AddNewDeviceTab: React.FC<IAddNewDeviceTabProps> = (props) => {
  const { service, devicesListName, onDeviceAdded } = props;
  const [title, setTitle] = React.useState('');
  const [deviceType, setDeviceType] = React.useState('');
  const [serialNumber, setSerialNumber] = React.useState('');
  const [isRequestable, setIsRequestable] = React.useState(true);
  const [submitting, setSubmitting] = React.useState(false);
  const [errorMsg, setErrorMsg] = React.useState<string | undefined>();
  const [successMsg, setSuccessMsg] = React.useState<string | undefined>();

  const handleSubmit = async () => {
    if (!title || !deviceType) {
      setErrorMsg('Title and Device Type are required.');
      return;
    }
    setSubmitting(true);
    setErrorMsg(undefined);
    setSuccessMsg(undefined);
    try {
      await service.addDevice(devicesListName, {
        Title: title,
        DeviceType: deviceType,
        SerialNumber: serialNumber,
        IsRequestable: isRequestable,
        Status: 'Available'
      });
      setSuccessMsg(`Device "${title}" added successfully!`);
      setTitle('');
      setDeviceType('');
      setSerialNumber('');
      setIsRequestable(true);
      await onDeviceAdded();
    } catch (e) {
      setErrorMsg(e instanceof Error ? e.message : 'Error adding device');
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { marginTop: 16, maxWidth: 500 } }}>
      {errorMsg && <MessageBar messageBarType={MessageBarType.error}>{errorMsg}</MessageBar>}
      {successMsg && <MessageBar messageBarType={MessageBarType.success}>{successMsg}</MessageBar>}

      <TextField label="Device Name / Model" value={title} onChange={(_, v) => setTitle(v || '')} required />
      <TextField label="Device Type" value={deviceType} onChange={(_, v) => setDeviceType(v || '')} required />
      <TextField label="Serial Number" value={serialNumber} onChange={(_, v) => setSerialNumber(v || '')} />
      <Checkbox label="Allow users to request this device?" checked={isRequestable} onChange={(_, checked) => setIsRequestable(!!checked)} />

      <div>
        <PrimaryButton text="Add Device" onClick={handleSubmit} disabled={submitting} />
      </div>
    </Stack>
  );
};
