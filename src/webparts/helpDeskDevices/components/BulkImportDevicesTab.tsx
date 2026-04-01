import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { HelpDeskDeviceService } from '../../../services/HelpDeskDeviceService';

export interface IBulkImportDevicesTabProps {
  service: HelpDeskDeviceService;
  devicesListName: string;
  onImportComplete: () => void;
}

export const BulkImportDevicesTab: React.FC<IBulkImportDevicesTabProps> = ({ service, devicesListName, onImportComplete }) => {
  const [file, setFile] = React.useState<File | null>(null);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);

  const onFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      setFile(e.target.files[0]);
    }
  };

  const handleImport = async () => {
    if (!file) return;

    setLoading(true);
    setError(null);

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const text = e.target?.result as string;
        const rows = text.split('\n').map(row => row.split(','));
        const headers = rows[0].map(h => h.trim().toLowerCase());
        const data = rows.slice(1).filter(r => r.length > 1);

        const devices = data.map(row => {
          return {
            Title: String(row[0] || '').trim(),
            Type: String(row[1] || '').trim(),
            Model: String(row[2] || '').trim(),
            SerialNumber: String(row[3] || '').trim(),
            Availability: true
          };
        }).filter(d => d.Title && d.Type);

        if (devices.length === 0) {
          throw new Error("No valid devices found in CSV.");
        }

        await service.addDeviceBulk(devicesListName, devices);
        onImportComplete();
      } catch (err: any) {
        setError(err.message || 'Error processing file');
        setLoading(false);
      }
    };
    reader.onerror = () => {
      setError('Error reading file');
      setLoading(false);
    };
    reader.readAsText(file);
  };

  return (
    <Stack tokens={{ childrenGap: 24 }} styles={{ root: { maxWidth: 600, marginTop: 16 } }}>
      <Stack tokens={{ childrenGap: 8 }}>
        <h3>Bulk Import Devices (CSV)</h3>
        <p>Ensure your CSV has columns: Title, Type, Model, SerialNumber</p>
        <input type="file" accept=".csv" onChange={onFileChange} />
        {file && <p>Selected: {file.name}</p>}
      </Stack>

      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}

      <Stack horizontal tokens={{ childrenGap: 12 }}>
        <PrimaryButton text="Import CSV" onClick={handleImport} disabled={!file || loading} />
        {loading && <Spinner size={SpinnerSize.small} />}
      </Stack>
    </Stack>
  );
};
