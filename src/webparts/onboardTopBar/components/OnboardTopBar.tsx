import * as React from 'react';
import styles from './OnboardTopBar.module.scss';
import { IOnboardTopBarProps } from './IOnboardTopBarProps';
import {
  ONBOARDING_CLUSTER_COLORS,
  IOnboardingClusterMetrics,
  ONBOARDING_CLUSTER_NAMES
} from '../../../services/onboarding/IOnboardingTaskMetrics';

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
  const parsedTotalTasks: number = Math.floor(Number(props.totalTasks));
  const totalTasks: number = Number.isFinite(parsedTotalTasks) ? Math.max(0, parsedTotalTasks) : 0;
  const completedTasks: number = Math.min(
    totalTasks,
    Math.max(0, Math.floor(Number(props.completedTasks) || 0))
  );
  const progress: number = totalTasks > 0 ? completedTasks / totalTasks : 0;
  const clusterProgress: IOnboardingClusterMetrics[] = props.clusterProgress ||
    ONBOARDING_CLUSTER_NAMES.map((name: string) => ({
      name,
      totalTasks: 1,
      completedTasks: 0
    }));
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
          {clusterProgress.map((cluster: IOnboardingClusterMetrics, index: number) => {
            const clusterCompletion: number = cluster.totalTasks > 0
              ? cluster.completedTasks / cluster.totalTasks
              : 0;
            const overallCompletion: number = progress * clusterProgress.length - index;
            const segmentProgress: number = props.clusterProgress
              ? Math.max(0, Math.min(1, clusterCompletion))
              : Math.max(0, Math.min(1, overallCompletion));
            const color: string = ONBOARDING_CLUSTER_COLORS[cluster.name] || '#64748b';

            return (
              <span
                className={styles.progressTrack}
                key={cluster.name}
                title={props.clusterProgress
                  ? `${cluster.name}: ${cluster.completedTasks}/${cluster.totalTasks}`
                  : undefined}
                style={{ flex: `${Math.max(cluster.totalTasks, 1)} 1 0` }}
              >
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
