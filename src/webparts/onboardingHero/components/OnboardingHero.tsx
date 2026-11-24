import * as React from 'react';
import styles from './OnboardingHero.module.scss';
import { IOnboardingHeroProps } from './IOnboardingHeroProps';
import OnboardTopBar from '../../onboardTopBar/components/OnboardTopBar';

function readInteger(value: number, fallback: number, minimum: number, maximum?: number): number {
  const parsedValue: number = Math.floor(Number(value));
  if (!Number.isFinite(parsedValue)) {
    return fallback;
  }

  return Math.max(minimum, Math.min(maximum === undefined ? parsedValue : maximum, parsedValue));
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

const OnboardingHero: React.FC<IOnboardingHeroProps> = (props: IOnboardingHeroProps) => {
  const [activeBanner, setActiveBanner] = React.useState<number>(0);
  const [isPaused, setIsPaused] = React.useState<boolean>(false);
  const totalDays: number = readInteger(props.totalDays, 100, 1);
  const currentDay: number = readInteger(props.currentDay, 1, 1, totalDays);
  const totalTasks: number = readInteger(props.totalTasks, 40, 1);
  const completedTasks: number = readInteger(props.completedTasks, 0, 0, totalTasks);
  const completedPercent: number = Math.floor((completedTasks / totalTasks) * 100);
  const openTaskCount: number = readInteger(props.openTaskCount, 0, 0);
  const learningMinutes: number = readInteger(props.learningMinutes, 0, 0);
  const userName: string = (props.userName || '').trim() || 'עובד/ת';
  const userRole: string = props.userRole || 'עובדת חדשה';
  const brandTitle: string = props.brandTitle || 'מסע הקליטה';
  const brandSubtitle: string = props.brandSubtitle || '100 ימים של קליטה';
  const bannerTitle: string = props.bannerTitle || '100 שנים של חזון · 100 ימים של קליטה';
  const mapUrl: string = props.mapUrl || '#journey-map';

  React.useEffect(() => {
    if (isPaused || window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
      return undefined;
    }

    const timer: number = window.setInterval(() => {
      setActiveBanner((current: number) => (current === 0 ? 1 : 0));
    }, 7000);

    return () => window.clearInterval(timer);
  }, [isPaused]);

  return (
    <React.Fragment>
      <OnboardTopBar
        brandTitle={brandTitle}
        brandSubtitle={brandSubtitle}
        displayName={userName}
        userRole={userRole}
        currentDay={currentDay}
        totalDays={totalDays}
        completedTasks={completedTasks}
        totalTasks={totalTasks}
      />
      <section className={styles.hero} dir="rtl" aria-label="פתיחת מסע הקליטה">
      <svg
        className={styles.cornerDecoration}
        width="180"
        height="140"
        viewBox="0 0 180 140"
        aria-hidden="true"
      >
        <path d="M88 0H104L0 104V88Z" fill="#A9C8EE" />
        <path d="M120 0H130L0 130V120Z" fill="#F2B33D" />
      </svg>

      <div className={styles.inner}>
        <div
          className={styles.banner}
          role="region"
          aria-label="שלט המסע"
          onMouseEnter={() => setIsPaused(true)}
          onMouseLeave={() => setIsPaused(false)}
        >
          <span className={styles.bannerLabel}><i aria-hidden="true" />מסע הקליטה</span>
          <div className={styles.bannerMessage} aria-live="polite" aria-atomic="true">
            {activeBanner === 0 ? (
              <strong key="title">{bannerTitle}</strong>
            ) : (
              <React.Fragment key="progress">
                <strong>יום {currentDay} מתוך {totalDays}</strong>
                <span className={styles.bannerProgress} aria-label={`${currentDay} מתוך ${totalDays} ימים`}>
                  <i style={{ width: `${(currentDay / totalDays) * 100}%` }} />
                </span>
              </React.Fragment>
            )}
          </div>
          <div className={styles.bannerDots} role="group" aria-label="בחירת הודעה">
            <button
              type="button"
              aria-label="מסר החזון"
              aria-pressed={activeBanner === 0}
              onFocus={() => setIsPaused(true)}
              onBlur={() => setIsPaused(false)}
              onClick={() => setActiveBanner(0)}
            />
            <button
              type="button"
              aria-label="ההתקדמות שלך"
              aria-pressed={activeBanner === 1}
              onFocus={() => setIsPaused(true)}
              onBlur={() => setIsPaused(false)}
              onClick={() => setActiveBanner(1)}
            />
          </div>
        </div>

        <div className={styles.heroMain}>
          <div className={styles.heroContent}>
            <div className={styles.greeting}>
              <span className={styles.avatar} aria-hidden="true">{getInitials(userName)}</span>
              <span>היי {userName},</span>
            </div>
            <h1>היום יום {currentDay} מתוך {totalDays}</h1>
            <div className={styles.welcome}>
              <p>
                ברוכה הבאה לנת״ע. החלום על רכבת קלה בתל אביב נולד ב-1925. ב-2023 נפתח הקו האדום,
                הקווים הירוק והסגול בבנייה, והמטרו בדרך. את מצטרפת בדיוק עכשיו, לדור שיבנה את ההמשך.
              </p>
              <p>
                גם המסע שלך מתחיל: {totalDays} ימי קליטה בשישה אשכולות תוכן. בכל יום ייפתחו לך רק
                המשימות שמתאימות לשלב שלך, כך שתדעי תמיד איפה את ומה מחכה לך.
              </p>
              <p className={styles.slogan}>כי היסטוריה עושים עם אנשים כמוך.</p>
            </div>

            <div className={styles.heroActions}>
              <a className={styles.mapButton} href={mapUrl}>
                <svg viewBox="0 0 24 24" aria-hidden="true">
                  <path d="M9 18 3 21V6l6-3 6 3 6-3v15l-6 3-6-3Z" />
                  <path d="M9 3v15M15 6v15" />
                </svg>
                <span>למפת המסע</span>
                <svg className={styles.downArrow} viewBox="0 0 24 24" aria-hidden="true">
                  <path d="M12 5v14m-7-7 7 7 7-7" />
                </svg>
              </a>

              <div className={styles.stats} aria-label="נתוני המסע">
                <div className={styles.stat}>
                  <strong>{completedPercent}%</strong>
                  <span>מהמשימות הסתיימו</span>
                </div>
                <div className={styles.stat}>
                  <strong>{openTaskCount}</strong>
                  <span>משימות פתוחות</span>
                </div>
                <div className={styles.stat}>
                  <strong>{learningMinutes} דק׳</strong>
                  <span>למידה פתוחה</span>
                </div>
              </div>
            </div>
          </div>

          <aside className={styles.story} aria-label="סיפור המסגרת">
            <span className={styles.storyEyebrow}>סיפור המסגרת</span>
            <h2>100 שנים של חזון</h2>
            <ol className={styles.timeline}>
              <li className={styles.past}>
                <strong>1925</strong>
                <span>נולד החלום על רכבת קלה בתל אביב</span>
              </li>
              <li className={styles.red}>
                <strong>2023</strong>
                <span>נפתח הקו האדום</span>
              </li>
              <li className={styles.now}>
                <span className={styles.timelineAvatar} aria-hidden="true">{getInitials(userName)}</span>
                <strong>היום</strong>
                <span>הקווים הירוק והסגול בבנייה, ואת מצטרפת</span>
              </li>
              <li className={styles.next}>
                <strong>הפרק הבא</strong>
                <span>המטרו</span>
              </li>
            </ol>
          </aside>
        </div>

        <div className={styles.trainTrack} aria-hidden="true">
          <span className={styles.rail} />
          <span className={styles.train} />
        </div>
      </div>
      </section>
    </React.Fragment>
  );
};

export default OnboardingHero;
