import * as React from 'react';
import styles from './TasksSection.module.scss';
import { ONBOARDING_CLUSTER_COLORS } from '../../../services/onboarding/IOnboardingTaskMetrics';
import { IOnboardingTask } from '../../../services/onboarding/IOnboardingTask';
import { classifyOnboardingTask } from '../../../services/onboarding/classifyOnboardingTask';
import { getOnboardingTaskButtonLabel } from '../../../services/onboarding/getOnboardingTaskButtonLabel';
import { ITasksSectionProps } from './ITasksSectionProps';

interface ITask extends IOnboardingTask {
  color: string;
}

interface ITaskActionToast {
  title: string;
  body: string;
  completionId?: number;
  url?: string;
}

const LockIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true">
    <rect x="5" y="11" width="14" height="10" rx="2" />
    <path d="M8 11V8a4 4 0 0 1 8 0v3" />
  </svg>
);

const ExternalIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true">
    <path d="M14 4h6v6M20 4l-9 9" />
    <path d="M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5" />
  </svg>
);

const UndoIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true">
    <path d="M9 14 4 9l5-5" />
    <path d="M4 9h9a7 7 0 0 1 0 14h-2" />
  </svg>
);

const TaskImage: React.FC<{ color: string }> = ({ color }: { color: string }) => (
  <span className={styles.taskImage} style={{ backgroundColor: color }}>
    <svg className={styles.imagePattern} viewBox="0 0 320 150" preserveAspectRatio="xMidYMid slice" aria-hidden="true">
      <g fill="none" stroke="currentColor" strokeWidth="7" strokeLinecap="round" strokeLinejoin="round" opacity=".2">
        <path d="M-10 40H120l70 70h150" />
        <path d="m-10 150 80-80h130" />
        <path d="M250-10v70l50 50v70" />
      </g>
      <g fill="currentColor" opacity=".3">
        <circle cx="120" cy="40" r="9" />
        <circle cx="190" cy="110" r="9" />
        <circle cx="300" cy="110" r="9" />
      </g>
    </svg>
    <span className={styles.taskImageIcon} style={{ color: '#fff' }}>
      <svg viewBox="0 0 24 24" aria-hidden="true">
        <path d="M7 3h7l5 5v13H7a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2Z" />
        <path d="M14 3v5h5m-9 5h5m-5 4h5" />
      </svg>
    </span>
  </span>
);

const TasksSection: React.FC<ITasksSectionProps> = (props: ITasksSectionProps) => {
  const [pendingTaskIds, setPendingTaskIds] = React.useState<number[]>([]);
  const pendingTaskIdsRef = React.useRef<Set<number>>(new Set<number>());
  const [toast, setToast] = React.useState<ITaskActionToast | undefined>();
  const [isUndoing, setIsUndoing] = React.useState<boolean>(false);
  React.useEffect(() => {
    if (!toast) {
      return undefined;
    }
    const timeout: number = window.setTimeout(() => setToast(undefined), 6500);
    return () => window.clearTimeout(timeout);
  }, [toast]);

  const currentDay: number = Number.isFinite(props.currentDay) && props.currentDay > 0
    ? Math.floor(props.currentDay)
    : 1;
  const today: Date = new Date();
  console.log('TasksSection: currentDay =', currentDay, 'today =', today.toISOString(), 'props.tasks =', props.tasks);
  const allTasks: ITask[] = props.tasks
    .map((task: IOnboardingTask) => ({
      ...task,
      color: ONBOARDING_CLUSTER_COLORS[task.clusterName] || '#64748b'
    }))
    .sort((first: ITask, second: ITask) =>
      first.journeyDay - second.journeyDay || first.title.localeCompare(second.title)
    );
  const openTasks: ITask[] = allTasks.filter(
    (task: ITask) => task.journeyDay <= currentDay
  );
  const upcomingTasks: ITask[] = allTasks
    .filter((task: ITask) => task.journeyDay > currentDay)
    .slice(0, 3);
  const lateCount: number = openTasks.filter(
    (task: ITask) => classifyOnboardingTask(task, currentDay, today) === 'overdue'
  ).length;
  const todayCount: number = openTasks.filter(
    (task: ITask) => task.journeyDay === currentDay
  ).length;

  const getSystemLocation = (url: string): string => {
    const lowerCaseUrl: string = url.toLowerCase();
    if (lowerCaseUrl.indexOf('luma') > -1) {
      return 'בלומה';
    }
    if (lowerCaseUrl.indexOf('sharepoint') > -1) {
      return 'ב-SharePoint';
    }
    if (lowerCaseUrl.indexOf('outlook') > -1) {
      return 'ביומן Outlook';
    }
    if (lowerCaseUrl.indexOf('whatsapp') > -1) {
      return 'בוואטסאפ';
    }
    return `ב-${new URL(url).hostname}`;
  };

  const handleTaskClick = async (task: ITask): Promise<void> => {
    if (pendingTaskIdsRef.current.has(task.id)) {
      return;
    }

    pendingTaskIdsRef.current.add(task.id);
    const reservedTab: Window | null = window.open('about:blank', '_blank');
    if (reservedTab) {
      reservedTab.opener = null;
    }
    setPendingTaskIds((currentIds: number[]) => currentIds.concat(task.id));

    try {
      const completionId: number = await props.onCompleteTask(task);
      setToast({
        title: task.taskContentType === 'מפגש'
          ? `פרטי המפגש והזימון נמצאים ${getSystemLocation(task.url)}`
          : `המשימה נפתחת ${getSystemLocation(task.url)} בלשונית חדשה`,
        body: 'המשימה סומנה כ״הסתיימה״.',
        completionId,
        url: reservedTab ? undefined : task.url
      });
      if (reservedTab && !reservedTab.closed) {
        reservedTab.location.replace(task.url);
      }
    } catch (error) {
      if (reservedTab && !reservedTab.closed) {
        reservedTab.close();
      }
      const message: string = error instanceof Error ? error.message : String(error);
      console.error('Could not complete the onboarding task.', error);
      setToast({
        title: 'לא ניתן לסמן את המשימה כהסתיימה',
        body: message
      });
    } finally {
      pendingTaskIdsRef.current.delete(task.id);
      setPendingTaskIds((currentIds: number[]) =>
        currentIds.filter((taskId: number) => taskId !== task.id)
      );
    }
  };

  const undoCompletion = async (): Promise<void> => {
    if (!toast || toast.completionId === undefined || isUndoing) {
      return;
    }
    setIsUndoing(true);
    try {
      await props.onUndoCompletion(toast.completionId);
      setToast({
        title: 'הפעולה בוטלה',
        body: 'המשימה חזרה לרשימת המשימות הפתוחות.'
      });
    } catch (error) {
      const message: string = error instanceof Error ? error.message : String(error);
      console.error('Could not undo the onboarding task completion.', error);
      setToast({
        title: 'לא ניתן לבטל את השלמת המשימה',
        body: message
      });
    } finally {
      setIsUndoing(false);
    }
  };

  const renderTask = (task: ITask, locked: boolean): React.ReactElement => {
    const status: string = classifyOnboardingTask(task, currentDay, today);
    const isLate: boolean = !locked && status === 'overdue';
    const isNewToday: boolean = !locked && status === 'newToday';
    const daysUntil: number = task.journeyDay - currentDay;
    const statusClass: string = locked
      ? styles.locked
      : (isLate ? styles.late : styles.today);
    const statusLabel: string = locked
      ? 'נעולה'
      : (isLate ? 'באיחור' : (isNewToday ? 'חדש היום' : ''));
    const durationLabel: string = task.duration === undefined
      ? `יום ${task.journeyDay}`
      : `${task.duration} דק׳ · יום ${task.journeyDay}`;
    const actionLabel: string = getOnboardingTaskButtonLabel(task.taskContentType);
    const isPending: boolean = pendingTaskIds.indexOf(task.id) > -1;

    return (
      <article className={`${styles.taskRow} ${statusClass}`} key={task.id}>
        <span className={styles.rail} aria-hidden="true" />
        <span className={styles.statusNode} aria-hidden="true">
          {locked ? (
            <LockIcon />
          ) : isLate ? (
            <svg viewBox="0 0 10 10" aria-hidden="true">
              <path d="M5 1.8v3.8" />
              <circle cx="5" cy="8" r="1" />
            </svg>
          ) : null}
        </span>
        <TaskImage color={task.color} />
        <div className={styles.taskInfo}>
          <div className={styles.taskMeta}>
            {statusLabel && (
              <span className={styles.statusPill}>
                {locked && <LockIcon />}
                {statusLabel}
              </span>
            )}
            <span className={styles.category}>
              <i style={{ backgroundColor: task.color }} />
              {task.clusterName}
            </span>
            <span>{durationLabel}</span>
          </div>
          <h4>{task.title}</h4>
        </div>
        {locked ? (
          <button type="button" className={styles.taskButton} disabled>
            {daysUntil === 1 ? 'נפתחת מחר' : `בעוד ${daysUntil} ימים`}
          </button>
        ) : (
          <button
            type="button"
            className={styles.taskButton}
            onClick={() => handleTaskClick(task)}
            disabled={isPending}
          >
            {isPending ? 'שומר...' : actionLabel}
            {!isPending && <ExternalIcon />}
          </button>
        )}
      </article>
    );
  };

  return (
    <section className={styles.tasksSection} dir="rtl" aria-label="המשימות שלך">
      <header className={styles.sectionHeader}>
        <div>
          <span className={styles.eyebrow}>יום {currentDay} · פתוח לביצוע</span>
          <h2>המשימות שלך</h2>
        </div>
        <div className={styles.summary} aria-label="סיכום המשימות">
          <span><i className={styles.lateDot} />{lateCount} באיחור</span>
          <span><i className={styles.todayDot} />{todayCount} של היום</span>
        </div>
      </header>

      {props.isLoading ? (
        <div className={styles.dataState}>טוען את המשימות שלך...</div>
      ) : props.loadError ? (
        <div className={`${styles.dataState} ${styles.dataError}`} role="alert">
          לא ניתן לטעון את המשימות: {props.loadError}
        </div>
      ) : (
        <div id="tasksSection" className={styles.columnsNew}>
          <section id="column1" className={styles.column1} aria-label="פתוח עכשיו">
            <h3>פתוח עכשיו <small>· {openTasks.length}</small></h3>
            <div className={styles.panel}>
              {openTasks.length > 0 ? (
                openTasks.map((task: ITask) => renderTask(task, false))
              ) : (
                <div className={styles.emptyState}>אין כרגע משימות פתוחות.</div>
              )}
            </div>
          </section>

          <section className={styles.column} aria-label="בקרוב במסע">
            <h3><LockIcon />בקרוב במסע <small>גלויות אך נעולות, וייפתחו כשיגיע המועד</small></h3>
            <div className={styles.panel}>
              {upcomingTasks.length > 0 ? (
                upcomingTasks.map((task: ITask) => renderTask(task, true))
              ) : (
                <div className={styles.emptyState}>אין משימות נוספות בקרוב.</div>
              )}
            </div>
          </section>
        </div>
      )}

      {toast && (
        <div className={styles.toastContainer} aria-live="polite">
          <div className={styles.toast} role="status">
            <span className={styles.toastIcon} aria-hidden="true">
              <ExternalIcon />
            </span>
            <div className={styles.toastText}>
              <strong>{toast.title}</strong>
              <span>{toast.body}</span>
            </div>
            {toast.url && (
              <a
                className={styles.toastOpenLink}
                href={toast.url}
                target="_blank"
                rel="noopener noreferrer"
              >
                פתיחת המשימה
              </a>
            )}
            {toast.completionId !== undefined && (
              <button
                type="button"
                className={styles.toastUndo}
                onClick={undoCompletion}
                disabled={isUndoing}
              >
                <UndoIcon />
                {isUndoing ? 'מבטל...' : 'ביטול'}
              </button>
            )}
          </div>
        </div>
      )}
    </section>
  );
};

export default TasksSection;
