import * as React from 'react';
import { useState, useEffect } from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';

import * as strings from 'ItSupportTicketsWebPartStrings';
import styles from './ItSupportTickets.module.scss';
import { ITopTicketService, ICreateTicketData } from '../../../services/ITopTicketService';

// ── Urgency options (static — matches iTop values) ────────────────────────────

const URGENCY_OPTIONS: IDropdownOption[] = [
  { key: '1', text: 'Critical' },
  { key: '2', text: 'High' },
  { key: '3', text: 'Medium' },
  { key: '4', text: 'Low' }
];

// ── Props ─────────────────────────────────────────────────────────────────────

export interface INewTicketPanelProps {
  service: ITopTicketService;
  currentUserEmail: string;
  onSubmitSuccess: () => void;
}

// ── Component ─────────────────────────────────────────────────────────────────

export const NewTicketPanel: React.FC<INewTicketPanelProps> = ({
  service,
  currentUserEmail,
  onSubmitSuccess
}) => {
  // ── Form state ──────────────────────────────────────────────────────────────
  const [title, setTitle] = useState('');
  const [description, setDescription] = useState('');
  const [serviceId, setServiceId] = useState<string>('');
  const [urgency, setUrgency] = useState<string>('3'); // default: Medium

  // ── Service dropdown data ───────────────────────────────────────────────────
  const [serviceOptions, setServiceOptions] = useState<IDropdownOption[]>([]);
  const [servicesLoading, setServicesLoading] = useState(true);
  const [servicesError, setServicesError] = useState<string | undefined>(undefined);

  // ── Submission state ────────────────────────────────────────────────────────
  const [submitting, setSubmitting] = useState(false);
  const [submitError, setSubmitError] = useState<string | undefined>(undefined);
  const [submitSuccess, setSubmitSuccess] = useState(false);

  // ── Validation ──────────────────────────────────────────────────────────────
  const [touched, setTouched] = useState(false);

  const titleError = touched && !title.trim() ? strings.ValidationTitleRequired : undefined;
  const descError = touched && !description.trim() ? strings.ValidationDescriptionRequired : undefined;
  const serviceError = touched && !serviceId ? strings.ValidationServiceRequired : undefined;

  // ── Load services on mount ──────────────────────────────────────────────────
  useEffect(() => {
    let cancelled = false;
    setServicesLoading(true);
    setServicesError(undefined);

    (async () => {
      try {
        const items = await service.getServices();
        if (!cancelled) {
          setServiceOptions(items.map((s) => ({ key: s.id, text: s.name })));
        }
      } catch (err) {
        if (!cancelled) setServicesError(strings.ErrorLoadServices);
        console.error('[NewTicketPanel] getServices error:', err);
      } finally {
        if (!cancelled) setServicesLoading(false);
      }
    })().catch(() => undefined);

    return () => { cancelled = true; };
  }, [service]);

  // ── Submit ──────────────────────────────────────────────────────────────────
  const handleSubmit = async (): Promise<void> => {
    setTouched(true);
    if (!title.trim() || !description.trim() || !serviceId) return;

    setSubmitting(true);
    setSubmitError(undefined);
    setSubmitSuccess(false);

    try {
      const data: ICreateTicketData = {
        title: title.trim(),
        description: description.trim(),
        serviceId,
        urgency,
        callerEmail: currentUserEmail
      };
      await service.createTicket(data);
      setSubmitSuccess(true);
      // Reset form
      setTitle('');
      setDescription('');
      setServiceId('');
      setUrgency('3');
      setTouched(false);
      // Notify parent to refresh My Tickets tab
      onSubmitSuccess();
    } catch (err) {
      setSubmitError(strings.ErrorSubmitTicket);
      console.error('[NewTicketPanel] createTicket error:', err);
    } finally {
      setSubmitting(false);
    }
  };

  const handleReset = (): void => {
    setTitle('');
    setDescription('');
    setServiceId('');
    setUrgency('3');
    setTouched(false);
    setSubmitError(undefined);
    setSubmitSuccess(false);
  };

  // ── Render ──────────────────────────────────────────────────────────────────
  return (
    <Stack tokens={{ childrenGap: 0 }} styles={{ root: { maxWidth: 640 } }}>
      <Text
        variant="mediumPlus"
        styles={{
          root: {
            fontWeight: 600,
            marginBottom: 20,
            color: 'var(--neutralPrimary, #323130)'
          }
        }}
      >
        {strings.FormTitle}
      </Text>

      {submitSuccess && (
        <MessageBar
          messageBarType={MessageBarType.success}
          onDismiss={() => setSubmitSuccess(false)}
          dismissButtonAriaLabel="Close"
          styles={{ root: { marginBottom: 16 } }}
        >
          {strings.SubmitSuccessMessage}
        </MessageBar>
      )}

      {submitError && (
        <MessageBar
          messageBarType={MessageBarType.error}
          onDismiss={() => setSubmitError(undefined)}
          dismissButtonAriaLabel="Close"
          styles={{ root: { marginBottom: 16 } }}
        >
          {submitError}
        </MessageBar>
      )}

      {servicesError && (
        <MessageBar
          messageBarType={MessageBarType.warning}
          styles={{ root: { marginBottom: 16 } }}
        >
          {servicesError}
        </MessageBar>
      )}

      <div className={styles.formField}>
        <TextField
          label={strings.FieldTitleLabel}
          placeholder={strings.FieldTitlePlaceholder}
          value={title}
          onChange={(_, v) => setTitle(v || '')}
          errorMessage={titleError}
          required
          disabled={submitting}
        />
      </div>

      <div className={styles.formField}>
        <TextField
          label={strings.FieldDescriptionLabel}
          placeholder={strings.FieldDescriptionPlaceholder}
          value={description}
          onChange={(_, v) => setDescription(v || '')}
          errorMessage={descError}
          multiline
          rows={5}
          required
          disabled={submitting}
        />
      </div>

      <div className={styles.formField}>
        <Dropdown
          label={strings.FieldServiceLabel}
          placeholder={
            servicesLoading ? 'Loading services…' : strings.FieldServicePlaceholder
          }
          options={serviceOptions}
          selectedKey={serviceId || undefined}
          onChange={(_, o) => setServiceId(String(o?.key ?? ''))}
          errorMessage={serviceError}
          required
          disabled={submitting || servicesLoading}
          styles={{ root: { maxWidth: 360 } }}
        />
      </div>

      <div className={styles.formField}>
        <Dropdown
          label={strings.FieldUrgencyLabel}
          options={URGENCY_OPTIONS}
          selectedKey={urgency}
          onChange={(_, o) => setUrgency(String(o?.key ?? '3'))}
          disabled={submitting}
          styles={{ root: { maxWidth: 200 } }}
        />
      </div>

      <div className={styles.formActions}>
        <PrimaryButton
          text={
            submitting
              ? undefined
              : strings.SubmitButtonLabel
          }
          onClick={handleSubmit}
          disabled={submitting}
          onRenderText={
            submitting
              ? () => (
                  <Stack horizontal tokens={{ childrenGap: 6 }} verticalAlign="center">
                    <Spinner size={SpinnerSize.xSmall} />
                    <span>Submitting…</span>
                  </Stack>
                )
              : undefined
          }
        />
        <DefaultButton
          text="Clear"
          onClick={handleReset}
          disabled={submitting}
        />
      </div>
    </Stack>
  );
};

export default NewTicketPanel;
