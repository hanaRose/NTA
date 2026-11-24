import * as React from 'react';
import styles from './OnboardTopBar.module.scss';
import { IOnboardTopBarProps } from './IOnboardTopBarProps';

const progressColors: string[] = [
  '#20B2A6',
  '#F2A832',
  '#4BAA43',
  '#9B4BCC',
  '#E8414F',
  '#4C78C9'
];

function toPositiveInteger(value: number, fallback: number): number {
  const parsedValue: number = Number(value);
  return Number.isFinite(parsedValue) && parsedValue > 0
    ? Math.floor(parsedValue)
    : fallback;
}

function getInitials(name: string): string {
  const words: string[] = name.trim().split(/\s+/).filter(Boolean);
  if (words.length === 0) {
    return '?';
  }

  return words.length > 1
    ? `${words[0].charAt(0)}${words[words.length - 1].charAt(0)}`
    : words[0].slice(0, 2);
}

const OnboardTopBar: React.FC<IOnboardTopBarProps> = (props: IOnboardTopBarProps) => {
  const totalDays: number = toPositiveInteger(props.totalDays, 100);
  const currentDay: number = Math.min(
    totalDays,
    Math.max(1, toPositiveInteger(props.currentDay, 1))
  );
  const totalTasks: number = toPositiveInteger(props.totalTasks, 1);
  const completedTasks: number = Math.min(
    totalTasks,
    Math.max(0, Math.floor(Number(props.completedTasks) || 0))
  );
  const progress: number = completedTasks / totalTasks;
  const displayName: string = (props.displayName || '').trim() || 'עובד/ת';
  const brandTitle: string = props.brandTitle || 'מסע הקליטה';
  const brandSubtitle: string = props.brandSubtitle || '100 ימים של קליטה';
  const userRole: string = props.userRole || 'עובד/ת חדשה';
  const topBarClass: string = props.standalone
    ? `${styles.onboardTopBar} ${styles.standalone}`
    : styles.onboardTopBar;

  return (
    <header className={topBarClass} dir="rtl">
      <div className={styles.brand}>
        <span className={styles.brandMark} aria-hidden="true">נת״ע</span>
        <span className={styles.brandText}>
          <strong>{brandTitle}</strong>
          <span>{brandSubtitle}</span>
        </span>
      </div>

      <section className={styles.progress} aria-label="התקדמות במסע הקליטה">
        <div className={styles.progressLabels}>
          <span>
            <strong>{completedTasks}</strong> מתוך {totalTasks} משימות הסתיימו
          </span>
          <span>
            יום <strong>{currentDay} / {totalDays}</strong>
          </span>
        </div>
        <div
          className={styles.progressTracks}
          role="img"
          aria-label={`הושלמו ${completedTasks} מתוך ${totalTasks} משימות`}
        >
          {progressColors.map((color: string, index: number) => {
            const segmentProgress: number = Math.max(
              0,
              Math.min(1, progress * progressColors.length - index)
            );

            return (
              <span className={styles.progressTrack} key={color}>
                <span
                  className={styles.progressFill}
                  style={{ width: `${segmentProgress * 100}%`, backgroundColor: color }}
                />
              </span>
            );
          })}
        </div>
      </section>

      <div className={styles.profile}>
        <span className={styles.avatar} aria-hidden="true">{getInitials(displayName)}</span>
        <span className={styles.profileText}>
          <strong>{displayName}</strong>
          <span>{userRole}</span>
        </span>
      </div>
    </header>
  );
};

export default OnboardTopBar;
