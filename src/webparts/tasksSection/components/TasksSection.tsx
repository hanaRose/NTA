import * as React from 'react';
import styles from './TasksSection.module.scss';
import { ITasksSectionProps } from './ITasksSectionProps';

type TaskIcon = 'video' | 'course' | 'document';

interface ITask {
  id: string;
  day: number;
  category: string;
  color: string;
  title: string;
  type: string;
  minutes: number;
  description: string;
  icon: TaskIcon;
  action: string;
  url?: string;
}

function readDay(value: number): number {
  const parsedValue: number = Math.floor(Number(value));
  return Number.isFinite(parsedValue) && parsedValue > 0 ? parsedValue : 14;
}

const ExternalIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true">
    <path d="M14 4h6v6M20 4l-9 9" />
    <path d="M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5" />
  </svg>
);

const LockIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true">
    <rect x="5" y="11" width="14" height="10" rx="2" />
    <path d="M8 11V8a4 4 0 0 1 8 0v3" />
  </svg>
);

const TaskImage: React.FC<{ task: ITask }> = ({ task }: { task: ITask }) => (
  <span className={styles.taskImage} style={{ backgroundColor: task.color }}>
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
    <span className={styles.taskImageIcon} style={{ color: task.id === 'safety-gear' ? '#232a5c' : '#fff' }}>
      {task.icon === 'video' && (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <circle cx="12" cy="12" r="9" />
          <path d="m10 8 6 4-6 4z" fill="currentColor" stroke="none" />
        </svg>
      )}
      {task.icon === 'course' && (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <rect x="3" y="4" width="18" height="13" rx="2" />
          <path d="M8 21h8m-4-4v4m-2-12 5 2-5 2z" />
        </svg>
      )}
      {task.icon === 'document' && (
        <svg viewBox="0 0 24 24" aria-hidden="true">
          <path d="M7 3h7l5 5v13H7a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2Z" />
          <path d="M14 3v5h5m-9 5h5m-5 4h5" />
        </svg>
      )}
    </span>
  </span>
);

const TasksSection: React.FC<ITasksSectionProps> = (props: ITasksSectionProps) => {
  const [selectedTask, setSelectedTask] = React.useState<ITask | undefined>();
  const currentDay: number = readDay(props.currentDay);
  const allTasks: ITask[] = [
    {
      id: 'payroll',
      day: 7,
      category: 'תנאי הנסיעה',
      color: '#4fa83d',
      title: 'נהלי שכר ונוכחות',
      type: 'סרטון',
      minutes: 6,
      description: 'מחלה, חופשה, הצהרה ובחירה: איך זה עובד אצלנו.',
      icon: 'video',
      action: 'לצפייה בסרטון',
      url: props.payrollUrl
    },
    {
      id: 'cyber',
      day: 10,
      category: 'חוקי הדרך',
      color: '#f2a93b',
      title: 'לומדת אבטחת מידע וסייבר',
      type: 'לומדה',
      minutes: 20,
      description: 'איך מזהים פישינג ושומרים על המידע של הפרויקטים.',
      icon: 'course',
      action: 'להתחלת הלומדה',
      url: props.cyberUrl
    },
    {
      id: 'safety-gear',
      day: 14,
      category: 'חוקי הדרך',
      color: '#f2a93b',
      title: 'קבלת ציוד בטיחות',
      type: 'מסמך אינטרקטיבי',
      minutes: 5,
      description: 'קסדה, וסט ונעלי בטיחות: מי זכאי ואיך מקבלים.',
      icon: 'document',
      action: 'לפתיחת המסמך',
      url: props.safetyGearUrl
    },
    {
      id: 'sap',
      day: 16,
      category: 'מרכז בקרה',
      color: '#9b4dca',
      title: 'הדרכת SAP',
      type: 'מסמך אינטרקטיבי',
      minutes: 12,
      description: 'הזמנות, אישורים ודיווחים: הצעדים הראשונים במערכת.',
      icon: 'document',
      action: 'לפתיחת המסמך',
      url: props.sapUrl
    },
    {
      id: 'bi',
      day: 19,
      category: 'מרכז בקרה',
      color: '#9b4dca',
      title: 'הדרכת BI',
      type: 'מסמך אינטרקטיבי',
      minutes: 10,
      description: 'הדשבורדים שמראים איך מתקדמים הפרויקטים.',
      icon: 'document',
      action: 'לפתיחת המסמך',
      url: props.biUrl
    },
    {
      id: 'knowledge',
      day: 21,
      category: 'מרכז בקרה',
      color: '#9b4dca',
      title: 'מערך ניהול הידע בארגון',
      type: 'מסמך אינטרקטיבי',
      minutes: 10,
      description: 'קבוצות עבודה מקצועיות, איגודים ואיך משתפים ידע.',
      icon: 'document',
      action: 'לפתיחת המסמך',
      url: props.knowledgeUrl
    }
  ];
  const openTasks: ITask[] = allTasks.filter((task: ITask) => task.day <= currentDay);
  const upcomingTasks: ITask[] = allTasks
    .filter((task: ITask) => task.day > currentDay)
    .slice(0, 3);
  const lateCount: number = openTasks.filter((task: ITask) => task.day < currentDay).length;
  const todayCount: number = openTasks.length - lateCount;

  React.useEffect(() => {
    if (!selectedTask) {
      return undefined;
    }

    const handleKeyDown = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        setSelectedTask(undefined);
      }
    };

    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [selectedTask]);

  const activateTask = (task: ITask): void => {
    if (task.url && task.url.trim()) {
      window.open(task.url, '_blank', 'noopener,noreferrer');
      return;
    }
    setSelectedTask(task);
  };

  const renderTask = (task: ITask, locked: boolean): React.ReactElement => {
    const isLate: boolean = !locked && task.day < currentDay;
    const daysUntil: number = task.day - currentDay;
    const statusClass: string = locked
      ? styles.locked
      : (isLate ? styles.late : styles.today);
    const statusLabel: string = locked
      ? 'נעולה'
      : (isLate ? 'באיחור' : 'חדש היום');
    const actionLabel: string = locked
      ? (daysUntil === 1 ? 'נפתחת מחר' : `בעוד ${daysUntil} ימים`)
      : task.action;

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
        <TaskImage task={task} />
        <div className={styles.taskInfo}>
          <div className={styles.taskMeta}>
            <span className={styles.statusPill}>
              {locked && <LockIcon />}
              {statusLabel}
            </span>
            <span className={styles.category}>
              <i style={{ backgroundColor: task.color }} />
              {task.category}
            </span>
            <span>{task.type} · {task.minutes} דק׳ · יום {task.day}</span>
          </div>
          <h4>{task.title}</h4>
          <p>{task.description}</p>
        </div>
        {locked ? (
          <button type="button" className={styles.taskButton} disabled>
            {actionLabel}
          </button>
        ) : (
          <button type="button" className={styles.taskButton} onClick={() => activateTask(task)}>
            {task.action}
            <ExternalIcon />
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

      <div className={styles.columns}>
        <section className={styles.column} aria-label="פתוח עכשיו">
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

      {selectedTask && (
        <div className={styles.backdrop} onClick={() => setSelectedTask(undefined)}>
          <section
            className={styles.dialog}
            role="dialog"
            aria-modal="true"
            aria-label={selectedTask.title}
            onClick={(event: React.MouseEvent<HTMLElement>) => event.stopPropagation()}
          >
            <button
              type="button"
              className={styles.closeButton}
              onClick={() => setSelectedTask(undefined)}
              aria-label="סגירה"
            >
              ×
            </button>
            <span className={styles.eyebrow}>{selectedTask.type} · יום {selectedTask.day}</span>
            <h2>{selectedTask.title}</h2>
            <p>{selectedTask.description}</p>
            <p className={styles.configureHint}>
              כדי לפתוח את התוכן, יש להגדיר את הקישור למשימה במאפייני הווב פרט.
            </p>
          </section>
        </div>
      )}
    </section>
  );
};

export default TasksSection;
