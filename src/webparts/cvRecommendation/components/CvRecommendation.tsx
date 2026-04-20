import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

import { ICvRecommendationProps } from './ICvRecommendationProps';
import {
  CvRecommendationService,
  ICvRecommendation,
  ICvStats
} from '../../../services/CvRecommendationService';
import CvDashboard from './CvDashboard';
import CvList from './CvList';
import CvSubmitForm from './CvSubmitForm';
import styles from './CvRecommendation.module.scss';

const CvRecommendation: React.FC<ICvRecommendationProps> = (props) => {
  const { title, accentColor, itemsPerPage, spHttpClient, siteUrl, currentUserId, currentUserName } = props;

  // ─── State ────────────────────────────────────────────────────────────
  const [items, setItems] = React.useState<ICvRecommendation[]>([]);
  const [stats, setStats] = React.useState<ICvStats>({ total: 0, submitted: 0, underReview: 0, accepted: 0, rejected: 0 });
  const [isHr, setIsHr] = React.useState<boolean>(false);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string>('');
  const [formOpen, setFormOpen] = React.useState<boolean>(false);

  // ─── Service ──────────────────────────────────────────────────────────
  const serviceRef = React.useRef<CvRecommendationService | null>(null);
  if (!serviceRef.current) {
    serviceRef.current = new CvRecommendationService(spHttpClient, siteUrl);
  }

  // ─── Init ─────────────────────────────────────────────────────────────
  const init = React.useCallback(async (): Promise<void> => {
    const svc = serviceRef.current;
    if (!svc) return;
    setLoading(true);
    setError('');
    try {
      await svc.ensureLists();
      const hrCheck = await svc.isHrUser();
      setIsHr(hrCheck);
      const data = await svc.getCvRecommendations(currentUserId, hrCheck);
      setItems(data);
      setStats(svc.computeStats(data));
    } catch (e) {
      console.error('CvRecommendation: init error', e);
      setError(e instanceof Error ? e.message : 'Failed to load CV recommendations.');
    } finally {
      setLoading(false);
    }
  }, [currentUserId]);

  React.useEffect(() => { init().catch(() => undefined); }, [init]);

  // ─── Refresh (after add/update/delete) ────────────────────────────────
  const handleRefresh = React.useCallback((): void => {
    init().catch(() => undefined);
  }, [init]);

  // ─── Submit new CV ────────────────────────────────────────────────────
  const handleSubmit = React.useCallback(async (
    candidateName: string,
    candidateEmail: string,
    phoneNumber: string,
    position: string,
    notes: string,
    file?: File
  ): Promise<void> => {
    const svc = serviceRef.current;
    if (!svc) throw new Error('Service not available.');
    await svc.submitCvRecommendation(candidateName, candidateEmail, phoneNumber, position, notes, file);
    handleRefresh();
  }, [handleRefresh]);

  // ─── Render ───────────────────────────────────────────────────────────
  return (
    <div className={styles.cvRecommendation}>

      {/* Header */}
      <div className={styles.header}>
        <div className={styles.headerLeft}>
          <Icon iconName="People" className={styles.titleIcon} />
          <h2 className={styles.title}>{title}</h2>
          <span className={`${styles.roleBadge} ${isHr ? styles.roleBadgeHr : styles.roleBadgeEmployee}`}>
            <Icon iconName={isHr ? 'Shield' : 'Contact'} style={{ fontSize: 10 }} />
            {isHr ? 'HR Admin' : 'Employee'}
          </span>
        </div>
        <PrimaryButton
          iconProps={{ iconName: 'Add' }}
          text="Recommend a CV"
          onClick={() => setFormOpen(true)}
          styles={{
            root: { backgroundColor: accentColor, borderColor: accentColor },
            rootHovered: { backgroundColor: accentColor, filter: 'brightness(0.9)' }
          }}
        />
      </div>

      {/* Error */}
      {error && (
        <MessageBar
          messageBarType={MessageBarType.error}
          onDismiss={() => setError('')}
          style={{ marginBottom: 16 }}
        >
          {error}
        </MessageBar>
      )}

      {/* Loading */}
      {loading ? (
        <div className={styles.stateBox}>
          <Spinner size={SpinnerSize.large} label="Loading CV recommendations…" />
        </div>
      ) : (
        <>
          {/* Dashboard / Stats */}
          <CvDashboard stats={stats} accentColor={accentColor} />

          {/* List */}
          {items.length === 0 && !error ? (
            <div className={styles.stateBox}>
              <Icon iconName="People" className={styles.stateIcon} />
              <div className={styles.stateTitle}>No CV Recommendations Yet</div>
              <div className={styles.stateDesc}>
                {isHr
                  ? 'No CV recommendations have been submitted yet.'
                  : `You haven't submitted any CV recommendations. Click "Recommend a CV" to get started.`}
              </div>
            </div>
          ) : (
            <CvList
              items={items}
              isHr={isHr}
              currentUserId={currentUserId}
              itemsPerPage={itemsPerPage}
              service={serviceRef.current!}
              onRefresh={handleRefresh}
              accentColor={accentColor}
            />
          )}
        </>
      )}

      {/* Submit Form Panel */}
      <CvSubmitForm
        isOpen={formOpen}
        onDismiss={() => setFormOpen(false)}
        onSubmit={handleSubmit}
      />

      {/* Current user context (debug/info) - hidden visually */}
      {currentUserName && (
        <div style={{ display: 'none' }} aria-hidden="true">
          signed-in: {currentUserName}
        </div>
      )}
    </div>
  );
};

export default CvRecommendation;
