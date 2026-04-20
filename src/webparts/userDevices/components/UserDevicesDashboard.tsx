import * as React from 'react';
import { useEffect, useState, useCallback, useRef } from 'react';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Stack } from '@fluentui/react/lib/Stack';
import { Icon } from '@fluentui/react/lib/Icon';

import { IUserDevicesProps } from './IUserDevicesProps';
import { AvailableUserDevicesTab } from './AvailableUserDevicesTab';
import { MyRequestedDevicesTab } from './MyRequestedDevicesTab';
import styles from './UserDevices.module.scss';
import { IHelpDeskDevice, IHelpDeskDeviceRequest } from '../../../services/HelpDeskDeviceService';

export const UserDevicesDashboard: React.FC<IUserDevicesProps> = (props) => {
  const { service, devicesListName, requestsListName } = props;

  const [currentUserId, setCurrentUserId] = useState<number | null>(null);
  const [initLoading, setInitLoading] = useState<boolean>(true);
  const [initError, setInitError] = useState<string | undefined>();

  const [devices, setDevices] = useState<IHelpDeskDevice[]>([]);
  const [requests, setRequests] = useState<IHelpDeskDeviceRequest[]>([]);
  const [dataLoading, setDataLoading] = useState<boolean>(false);
  const mountedRef = useRef(true);

  // 1. Get current user ID
  useEffect(() => {
    mountedRef.current = true;
    let cancelled = false;
    (async () => {
      setInitLoading(true);
      try {
        const id = await service.getCurrentUserId();
        if (!cancelled) {
          setCurrentUserId(id);
          setInitError(undefined);
        }
      } catch {
        if (!cancelled) setInitError('Unable to get user info.');
      } finally {
        if (!cancelled) setInitLoading(false);
      }
    })().catch(console.error);
    return () => { cancelled = true; mountedRef.current = false; };
  }, [service]);

  // 2. Load Data
  const loadData = useCallback(async () => {
    if (!currentUserId) return;
    setDataLoading(true);
    try {
      const [allDevices, myRequests] = await Promise.all([
        service.getDevices(devicesListName),
        service.getMyRequests(requestsListName, currentUserId)
      ]);
      if (mountedRef.current) {
        setDevices(allDevices);
        setRequests(myRequests);
      }
    } catch (err) {
      if (mountedRef.current) console.error('Error loading data:', err);
    } finally {
      if (mountedRef.current) setDataLoading(false);
    }
  }, [service, devicesListName, requestsListName, currentUserId]);

  useEffect(() => {
    if (!initLoading && currentUserId) {
      loadData().catch(console.error);
    }
  }, [initLoading, currentUserId, loadData]);

  if (initLoading) {
    return (
      <div className={styles.userDevices}>
        <div className={styles.wpHeader}>
          <div className={styles.wpTitleSection}>
            <div className={styles.wpIconWrap}><Icon iconName="Devices2" /></div>
            <h2 className={styles.wpTitle}>Devices Catalog</h2>
          </div>
        </div>
        <Stack horizontalAlign="center" styles={{ root: { padding: 32 } }}>
          <Spinner size={SpinnerSize.large} label="Loading profile..." />
        </Stack>
      </div>
    );
  }

  if (initError || !currentUserId) {
    return (
      <div className={styles.userDevices}>
        <div className={styles.wpHeader}>
          <div className={styles.wpTitleSection}>
            <div className={styles.wpIconWrap}><Icon iconName="Devices2" /></div>
            <h2 className={styles.wpTitle}>Devices Catalog</h2>
          </div>
        </div>
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
          {initError || 'User not found'}
        </MessageBar>
      </div>
    );
  }

  return (
    <div className={styles.userDevices}>
      <div className={styles.wpHeader}>
        <div className={styles.wpTitleSection}>
          <div className={styles.wpIconWrap}><Icon iconName="Devices2" /></div>
          <h2 className={styles.wpTitle}>Devices Catalog</h2>
        </div>
      </div>

      <Pivot>
        <PivotItem headerText="Available Devices" itemIcon="Devices3">
          <AvailableUserDevicesTab
            devices={devices}
            service={service}
            requestsListName={requestsListName}
            currentUserId={currentUserId}
            onDataChange={loadData}
            loading={dataLoading}
          />
        </PivotItem>
        <PivotItem headerText="My Requested Devices" itemIcon="PageList">
          <MyRequestedDevicesTab
            requests={requests}
            devices={devices}
            loading={dataLoading}
          />
        </PivotItem>
      </Pivot>
    </div>
  );
};


