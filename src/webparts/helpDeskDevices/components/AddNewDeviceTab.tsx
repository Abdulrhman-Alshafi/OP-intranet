import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { TextField } from '@fluentui/react/lib/TextField';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Icon } from '@fluentui/react/lib/Icon';
import { HelpDeskDeviceService } from '../../../services/HelpDeskDeviceService';
import styles from './HelpDeskDevices.module.scss';

const STATUS_OPTIONS: IDropdownOption[] = [
  { key: 'Available', text: 'Available' },
  { key: 'In Use', text: 'In Use' },
  { key: 'Under Maintenance', text: 'Under Maintenance' },
  { key: 'Retired', text: 'Retired' },
];

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
  const [status, setStatus] = React.useState<string>('Available');
  const [isRequestable, setIsRequestable] = React.useState(true);
  const [submitting, setSubmitting] = React.useState(false);
  const [errorMsg, setErrorMsg] = React.useState<string | undefined>();
  const [successMsg, setSuccessMsg] = React.useState<string | undefined>();

  const handleSubmit = async () => {
    if (!title || !deviceType) {
      setErrorMsg('Device Name and Device Type are required.');
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
        Status: status
      });
      setSuccessMsg(`Device "${title}" added successfully!`);
      setTitle('');
      setDeviceType('');
      setSerialNumber('');
      setStatus('Available');
      setIsRequestable(true);
      await onDeviceAdded();
    } catch (e) {
      setErrorMsg(e instanceof Error ? e.message : 'Error adding device');
    } finally {
      setSubmitting(false);
    }
  };

  const handleReset = () => {
    setTitle('');
    setDeviceType('');
    setSerialNumber('');
    setStatus('Available');
    setIsRequestable(true);
    setErrorMsg(undefined);
    setSuccessMsg(undefined);
  };

  return (
    <div style={{ paddingTop: 16 }}>
      <h3 className={styles.uploadSectionTitle}>
        <Icon iconName="Add" />
        Add a Single Device
      </h3>

      <Stack tokens={{ childrenGap: 14 }} styles={{ root: { maxWidth: 480 } }}>
        {errorMsg && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setErrorMsg(undefined)}>
            {errorMsg}
          </MessageBar>
        )}
        {successMsg && (
          <MessageBar messageBarType={MessageBarType.success} onDismiss={() => setSuccessMsg(undefined)}>
            {successMsg}
          </MessageBar>
        )}

        <div className={styles.uploadField}>
          <TextField
            label="Device Name / Model"
            value={title}
            onChange={(_, v) => setTitle(v || '')}
            required
            placeholder="e.g. Dell Latitude 5540"
          />
        </div>
        <div className={styles.uploadField}>
          <TextField
            label="Device Type"
            value={deviceType}
            onChange={(_, v) => setDeviceType(v || '')}
            required
            placeholder="e.g. Laptop, Desktop, Tablet"
          />
        </div>
        <div className={styles.uploadField}>
          <TextField
            label="Serial Number"
            value={serialNumber}
            onChange={(_, v) => setSerialNumber(v || '')}
            placeholder="Optional"
          />
        </div>
        <div className={styles.uploadField}>
          <Dropdown
            label="Status"
            selectedKey={status}
            options={STATUS_OPTIONS}
            onChange={(_, o) => setStatus(o?.key as string ?? 'Available')}
          />
        </div>
        <div className={styles.uploadField}>
          <Checkbox
            label="Allow users to request this device?"
            checked={isRequestable}
            onChange={(_, checked) => setIsRequestable(!!checked)}
          />
        </div>

        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text={submitting ? 'Adding...' : 'Add Device'}
            iconProps={{ iconName: 'Add' }}
            onClick={handleSubmit}
            disabled={submitting}
          />
          <DefaultButton text="Reset" onClick={handleReset} disabled={submitting} />
        </Stack>
      </Stack>
    </div>
  );
};


export interface IAddNewDeviceTabProps {
  service: HelpDeskDeviceService;
  devicesListName: string;
  onDeviceAdded: () => Promise<void>;
}

