import * as React from 'react';
import importedStyles from './CountdownTimer.module.scss';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const styles: any = importedStyles;
import type { ICountdownTimerProps } from './ICountdownTimerProps';
import { escape } from '@microsoft/sp-lodash-subset';

interface ITimeLeft {
  days: number;
  hours: number;
  minutes: number;
  seconds: number;
}

const CountdownTimer = (props: ICountdownTimerProps): React.ReactElement => {
  const {
    eventTitle,
    eventDescription,
    targetDate,
    backgroundImage,
    overlayOpacity,
    layout,
    showDays,
    showHours,
    showMinutes,
    showSeconds,
    completedMessage,
    accentColor
  } = props;

  const calculateTimeLeft = React.useCallback((): ITimeLeft | null => {
    if (!targetDate) {
      return null;
    }

    const difference = new Date(targetDate).getTime() - new Date().getTime();
    if (difference <= 0) {
      return null;
    }

    return {
      days: Math.floor(difference / (1000 * 60 * 60 * 24)),
      hours: Math.floor((difference / (1000 * 60 * 60)) % 24),
      minutes: Math.floor((difference / 1000 / 60) % 60),
      seconds: Math.floor((difference / 1000) % 60)
    };
  }, [targetDate]);

  const [timeLeft, setTimeLeft] = React.useState<ITimeLeft | null>(() => calculateTimeLeft());

  React.useEffect(() => {
    setTimeLeft(calculateTimeLeft());

    const interval = window.setInterval(() => {
      setTimeLeft(calculateTimeLeft());
    }, 1000);

    return () => {
      window.clearInterval(interval);
    };
  }, [calculateTimeLeft]);

  const hasImage = !!backgroundImage;
  const accent = accentColor || '#0078d4';
  const isCompact = layout === 'compact';
  const isBanner = layout === 'banner';

  if (!targetDate) {
    return (
      <section className={styles.countdownContainer}>
        <div className={styles.emptyState}>
          <h2>Countdown Timer</h2>
          <p>Set a target date in the property pane to start the countdown.</p>
        </div>
      </section>
    );
  }

  const units: { value: number; label: string }[] = [];
  if (timeLeft) {
    if (showDays) units.push({ value: timeLeft.days, label: 'Days' });
    if (showHours) units.push({ value: timeLeft.hours, label: 'Hours' });
    if (showMinutes) units.push({ value: timeLeft.minutes, label: 'Minutes' });
    if (showSeconds) units.push({ value: timeLeft.seconds, label: 'Seconds' });
  }

  const titleClass = `${styles.eventTitle} ${hasImage ? styles.eventTitleOnImage : ''}`;
  const descClass = `${styles.eventDescription} ${hasImage ? styles.eventDescriptionOnImage : ''}`;

  const timerContent = timeLeft ? (
    <div className={`${styles.timerGrid} ${isCompact ? styles.timerGridCompact : ''}`}>
      {units.map((unit, idx) => (
        <React.Fragment key={unit.label}>
          {idx > 0 && (
            <div className={`${styles.separator} ${hasImage ? styles.separatorOnImage : ''}`}>:</div>
          )}
          <div className={`${styles.timeUnit} ${hasImage ? styles.timeUnitOnImage : ''} ${isCompact ? styles.timeUnitCompact : ''}`}>
            <div className={`${styles.accentBar} ${isCompact ? styles.accentBarCompact : ''}`} style={{ backgroundColor: accent }} />
            <span className={`${styles.timeValue} ${hasImage ? styles.timeValueOnImage : ''} ${isCompact ? styles.timeValueCompact : ''}`}>
              {unit.value < 10 ? '0' + unit.value : '' + unit.value}
            </span>
            <span className={`${styles.timeLabel} ${hasImage ? styles.timeLabelOnImage : ''}`}>
              {unit.label}
            </span>
          </div>
        </React.Fragment>
      ))}
    </div>
  ) : (
    <div className={`${styles.completedMessage} ${hasImage ? styles.completedMessageOnImage : ''}`}>
      <span className={styles.checkIcon} style={{ backgroundColor: accent }}>✓</span>
      {escape(completedMessage || 'The event has started!')}
    </div>
  );

  return (
    <section className={styles.countdownContainer}>
      {hasImage && (
        <>
          <div className={styles.bgImage} style={{ backgroundImage: `url(${backgroundImage})` }} />
          <div className={styles.bgOverlay} style={{ backgroundColor: `rgba(0, 0, 0, ${(overlayOpacity || 50) / 100})` }} />
        </>
      )}

      <div className={styles.content}>
        {isBanner ? (
          <div className={styles.bannerLayout}>
            <div className={styles.bannerLeft}>
              {eventTitle && <h2 className={titleClass}>{escape(eventTitle)}</h2>}
              {eventDescription && <p className={descClass}>{escape(eventDescription)}</p>}
            </div>
            <div className={styles.bannerRight}>
              {timerContent}
            </div>
          </div>
        ) : (
          <>
            {eventTitle && <h2 className={titleClass}>{escape(eventTitle)}</h2>}
            {eventDescription && <p className={descClass}>{escape(eventDescription)}</p>}
            {timerContent}
          </>
        )}
      </div>
    </section>
  );
};

export default CountdownTimer;
