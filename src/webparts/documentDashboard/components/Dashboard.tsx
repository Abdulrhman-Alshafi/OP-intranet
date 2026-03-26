import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { getTheme, mergeStyleSets } from '@fluentui/react/lib/Styling';
import { IDepartmentDocuments, ISharePointDocument, SharePointService } from '../../../services/sharepointService';
import { IDashboardProps } from './IDashboardProps';
import PersonalDocuments from './PersonalDocuments';
import DepartmentDocuments from './DepartmentDocuments';

const Dashboard: React.FC<IDashboardProps> = ({
  title,
  libraryName,
  userEmail,
  siteUrl,
  webServerRelativeUrl,
  spHttpClient,
  themeVariant
}) => {
  const [personalFiles, setPersonalFiles] = useState<ISharePointDocument[]>([]);
  const [departmentFiles, setDepartmentFiles] = useState<IDepartmentDocuments[]>([]);

  const [loadingPersonal, setLoadingPersonal] = useState<boolean>(true);
  const [loadingDepartment, setLoadingDepartment] = useState<boolean>(true);

  const [personalError, setPersonalError] = useState<string | undefined>(undefined);
  const [departmentError, setDepartmentError] = useState<string | undefined>(undefined);
  const [searchQuery, setSearchQuery] = useState<string>('');

  const sharePointService = useMemo(
    () => new SharePointService(siteUrl, webServerRelativeUrl, spHttpClient),
    [siteUrl, webServerRelativeUrl, spHttpClient]
  );

  useEffect(() => {
    let cancelled = false;

    const fetchPersonalDocuments = async (): Promise<void> => {
      setLoadingPersonal(true);
      setPersonalError(undefined);

      try {
        const files = await sharePointService.getPersonalFiles(libraryName, userEmail);
        if (!cancelled) {
          setPersonalFiles(files);
        }
      } catch (error) {
        if (!cancelled) {
          setPersonalError(
            error instanceof Error ? error.message : 'Unable to load personal documents.'
          );
        }
      } finally {
        if (!cancelled) {
          setLoadingPersonal(false);
        }
      }
    };

    const fetchDepartmentDocuments = async (): Promise<void> => {
      setLoadingDepartment(true);
      setDepartmentError(undefined);

      try {
        const filesByDepartment = await sharePointService.getDepartmentFiles(libraryName);
        if (!cancelled) {
          setDepartmentFiles(filesByDepartment);
        }
      } catch (error) {
        if (!cancelled) {
          setDepartmentError(
            error instanceof Error ? error.message : 'Unable to load department documents.'
          );
        }
      } finally {
        if (!cancelled) {
          setLoadingDepartment(false);
        }
      }
    };

    fetchPersonalDocuments().catch(() => undefined);
    fetchDepartmentDocuments().catch(() => undefined);

    return () => {
      cancelled = true;
    };
  }, [sharePointService, libraryName, userEmail]);

  const filteredPersonalFiles = useMemo(() => {
    if (!searchQuery.trim()) {
      return personalFiles;
    }

    const query = searchQuery.toLowerCase();
    return personalFiles.filter((file) => file.name.toLowerCase().includes(query));
  }, [personalFiles, searchQuery]);

  const filteredDepartmentFiles = useMemo(() => {
    if (!searchQuery.trim()) {
      return departmentFiles;
    }

    const query = searchQuery.toLowerCase();
    return departmentFiles
      .map((department) => ({
        ...department,
        files: department.files.filter((file) => file.name.toLowerCase().includes(query))
      }))
      .filter((department) => department.files.length > 0);
  }, [departmentFiles, searchQuery]);

  const baseTheme = getTheme();
  const semanticColors = themeVariant?.semanticColors;

  const classNames = mergeStyleSets({
    root: {
      backgroundColor: semanticColors?.bodyBackground || baseTheme.palette.white,
      color: semanticColors?.bodyText || baseTheme.palette.neutralPrimary,
      padding: 16,
      borderRadius: 2
    },
    sectionTitle: {
      fontWeight: 600
    }
  });

  return (
    <Stack tokens={{ childrenGap: 16 }} className={classNames.root}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="end" wrap>
        <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
          {title}
        </Text>
        <SearchBox
          placeholder="Search documents"
          value={searchQuery}
          onChange={(_, value) => setSearchQuery(value || '')}
          styles={{ root: { width: 320, maxWidth: '100%' } }}
        />
      </Stack>

      <Pivot>
        <PivotItem
          headerText="My Documents"
          itemKey="myDocuments"
          itemCount={filteredPersonalFiles.length}
        >
          <Stack tokens={{ childrenGap: 8 }} styles={{ root: { paddingTop: 12 } }}>
            <Text variant="mediumPlus" className={classNames.sectionTitle}>
              My Documents
            </Text>
            <PersonalDocuments
              files={filteredPersonalFiles}
              loading={loadingPersonal}
              error={personalError}
              siteUrl={siteUrl}
            />
          </Stack>
        </PivotItem>

        <PivotItem
          headerText="Department Documents"
          itemKey="departmentDocuments"
          itemCount={filteredDepartmentFiles.reduce(
            (total, department) => total + department.files.length,
            0
          )}
        >
          <Stack tokens={{ childrenGap: 8 }} styles={{ root: { paddingTop: 12 } }}>
            <Text variant="mediumPlus" className={classNames.sectionTitle}>
              Department Documents
            </Text>
            <DepartmentDocuments
              departmentFiles={filteredDepartmentFiles}
              loading={loadingDepartment}
              error={departmentError}
              siteUrl={siteUrl}
            />
          </Stack>
        </PivotItem>
      </Pivot>
    </Stack>
  );
};

export default Dashboard;
