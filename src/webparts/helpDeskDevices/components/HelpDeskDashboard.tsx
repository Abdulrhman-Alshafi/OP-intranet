import * as React from 'react';
import { useEffect, useState, useCallback, useRef } from 'react';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Stack } from '@fluentui/react/lib/Stack';
import { Icon } from '@fluentui/react/lib/Icon';

import { IHelpDeskDevicesProps } from './IHelpDeskDevicesProps';
import { RequestedDevicesTab } from './RequestedDevicesTab';
import { AvailableDevicesTab } from './AvailableDevicesTab';
import { AddNewDeviceTab } from './AddNewDeviceTab';
import { BulkImportDevicesTab } from './BulkImportDevicesTab';
import styles from './HelpDeskDevices.module.scss';
import { IHelpDeskDevice, IHelpDeskDeviceRequest } from '../../../services/HelpDeskDeviceService';

export const HelpDeskDashboard: React.FC<IHelpDeskDevicesProps> = (props) => {
  const { service, devicesListName, requestsListName, helpDeskGroupName } = props;

  const [hasAccess, setHasAccess] = useState<boolean>(false);
  const [authLoading, setAuthLoading] = useState<boolean>(true);
  const [authError, setAuthError] = useState<string | undefined>();

  const [devices, setDevices] = useState<IHelpDeskDevice[]>([]);
  const [requests, setRequests] = useState<IHelpDeskDeviceRequest[]>([]);
  const [dataLoading, setDataLoading] = useState<boolean>(false);
  const mountedRef = useRef(true);

  // 1. Auth check
  useEffect(() => {
    mountedRef.current = true;
    let cancelled = false;
    (async () => {
      setAuthLoading(true);
      try {
        const [isGroupMember, isSiteAdmin] = await Promise.all([
          service.isUserInGroup(helpDeskGroupName),
          service.isCurrentUserSiteAdmin()
        ]);
        if (!cancelled) {
          setHasAccess(isGroupMember || isSiteAdmin);
          setAuthError(undefined);
        }
      } catch (err) {
        if (!cancelled) setAuthError('Unable to verify permissions.');
      } finally {
        if (!cancelled) setAuthLoading(false);
      }
    })().catch(console.error);
    return () => { cancelled = true; mountedRef.current = false; };
  }, [service, helpDeskGroupName]);

  // 2. Load Data
  const loadData = useCallback(async () => {
    setDataLoading(true);
    try {
      const [allDevices, allRequests] = await Promise.all([
        service.getDevices(devicesListName),
        service.getRequests(requestsListName)
      ]);
      if (mountedRef.current) {
        setDevices(allDevices);
        setRequests(allRequests);
      }
    } catch (err) {
      if (mountedRef.current) console.error('Error loading data:', err);
    } finally {
      if (mountedRef.current) setDataLoading(false);
    }
  }, [service, devicesListName, requestsListName]);

  useEffect(() => {
    if (!authLoading && hasAccess) {
      loadData().catch(console.error);
    }
  }, [authLoading, hasAccess, loadData]);

  if (authLoading) {
    return (
      <div className={styles.helpDeskDevices}>
        <div className={styles.wpHeader}>
          <div className={styles.wpTitleSection}>
            <div className={styles.wpIconWrap}><Icon iconName="Devices3" /></div>
            <h2 className={styles.wpTitle}>Help Desk Devices Admin</h2>
          </div>
        </div>
        <Stack horizontalAlign="center" styles={{ root: { padding: 32 } }}>
          <Spinner size={SpinnerSize.large} label="Verifying access..." />
        </Stack>
      </div>
    );
  }

  if (authError) {
    return (
      <div className={styles.helpDeskDevices}>
        <div className={styles.wpHeader}>
          <div className={styles.wpTitleSection}>
            <div className={styles.wpIconWrap}><Icon iconName="Devices3" /></div>
            <h2 className={styles.wpTitle}>Help Desk Devices Admin</h2>
          </div>
        </div>
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
          {authError}
        </MessageBar>
      </div>
    );
  }

  if (!hasAccess) {
    return (
      <div className={styles.helpDeskDevices}>
        <div className={styles.accessDenied}>
          <Icon iconName="Lock" styles={{ root: { fontSize: 40, marginBottom: 12, color: '#a4262c' } }} />
          <h2 style={{ margin: '0 0 8px' }}>Access Denied</h2>
          <p style={{ margin: 0, color: '#605e5c' }}>You do not have permission to view the Help Desk Devices dashboard.</p>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.helpDeskDevices}>
      <div className={styles.wpHeader}>
        <div className={styles.wpTitleSection}>
          <div className={styles.wpIconWrap}><Icon iconName="Devices3" /></div>
          <h2 className={styles.wpTitle}>Help Desk Devices Admin</h2>
        </div>
      </div>

      <Pivot>
        <PivotItem headerText="Requested Devices" itemIcon="InboxCheck">
          <RequestedDevicesTab
            requests={requests}
            devices={devices}
            service={service}
            requestsListName={requestsListName}
            onDataChange={loadData}
            loading={dataLoading}
          />
        </PivotItem>
        <PivotItem headerText="Available Devices" itemIcon="Devices3">
          <AvailableDevicesTab
            devices={devices}
            service={service}
            devicesListName={devicesListName}
            onDataChange={loadData}
            loading={dataLoading}
          />
        </PivotItem>
        <PivotItem headerText="Add Device" itemIcon="Add">
          <AddNewDeviceTab
            service={service}
            devicesListName={devicesListName}
            onDeviceAdded={loadData}
          />
        </PivotItem>
        <PivotItem headerText="Bulk Import" itemIcon="BulkUpload">
          <BulkImportDevicesTab
            service={service}
            devicesListName={devicesListName}
            onImportComplete={loadData}
          />
        </PivotItem>
      </Pivot>
    </div>
  );
};

