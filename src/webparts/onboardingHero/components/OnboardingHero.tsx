import * as React from 'react';
import styles from './OnboardingHero.module.scss';
import { HeroTimelineStepVariant, IHeroTimelineStep, IOnboardingHeroProps } from './IOnboardingHeroProps';
import OnboardTopBar from '../../onboardTopBar/components/OnboardTopBar';

const HEEBO_FONT_STYLESHEET: string = 'https://fonts.googleapis.com/css2?family=Heebo:wght@400;500;600;700;800;900&display=swap';

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

function getTimelineClass(variant: HeroTimelineStepVariant): string {
  switch (variant) {
    case 'past':
      return styles.past;
    case 'red':
      return styles.red;
    case 'now':
      return styles.now;
    case 'next':
      return styles.next;
    default:
      return styles.past;
  }
}

const OnboardingHero: React.FC<IOnboardingHeroProps> = (props: IOnboardingHeroProps) => {
  const [bannerState, setBannerState] = React.useState<{ active: number; leaving: number | null }>({
    active: 0,
    leaving: null
  });
  const [isPaused, setIsPaused] = React.useState<boolean>(false);
  const totalDays: number = readInteger(props.totalDays, 100, 1);
  const currentDay: number = readInteger(props.currentDay, 1, 1, totalDays);
  const totalTasks: number = readInteger(props.totalTasks, 0, 0);
  const completedTasks: number = readInteger(props.completedTasks, 0, 0, totalTasks);
  const completedPercent: number = totalTasks > 0
    ? Math.floor((completedTasks / totalTasks) * 100)
    : 0;
  const openTaskCount: number = readInteger(props.openTaskCount, 0, 0);
  const learningMinutes: number = readInteger(props.learningMinutes, 0, 0);
  const userName: string = (props.userName || '').trim() || 'עובד/ת';
  const userRole: string = props.userRole || 'עובדת חדשה';
  const brandTitle: string = props.brandTitle || 'מסע הקליטה';
  const brandSubtitle: string = props.brandSubtitle || '100 ימים של קליטה';
  const bannerTitle: string = props.bannerTitle || '100 שנים של חזון · 100 ימים של קליטה';
  const mapUrl: string = props.mapUrl || '#journeyMap';
  const welcomeParagraph1: string = props.welcomeParagraph1 || 'ברוכה הבאה לנת״ע. החלום על רכבת קלה בתל אביב נולד ב-1925. ב-2023 נפתח הקו האדום, הקווים הירוק והסגול בבנייה, והמטרו בדרך. את מצטרפת בדיוק עכשיו, לדור שיבנה את ההמשך.';
  const welcomeParagraph2: string = props.welcomeParagraph2 || 'גם המסע שלך מתחיל: 100 ימי קליטה בשישה אשכולות תוכן. בכל יום ייפתחו לך רק המשימות שמתאימות לשלב שלך, כך שתדעי תמיד איפה את ומה מחכה לך.';
  const slogan: string = props.slogan || 'כי היסטוריה עושים עם אנשים כמוך.';
  const storyTitle: string = props.storyTitle || '100 שנים של חזון';
  const activeBanner: number = bannerState.active;

  const showBanner = (index: number): void => {
    setBannerState((current) => current.active === index
      ? current
      : { active: index, leaving: current.active });
  };

  const getBannerMessageClass = (index: number): string => {
    if (bannerState.active === index) {
      return styles.bannerMessageActive;
    }
    if (bannerState.leaving === index) {
      return styles.bannerMessageLeaving;
    }
    return styles.bannerMessageHidden;
  };

  React.useEffect(() => {
    const hasHeeboStylesheet: boolean = Array.from(
      document.head.querySelectorAll<HTMLLinkElement>('link[rel="stylesheet"]')
    ).some((link: HTMLLinkElement) => link.href === HEEBO_FONT_STYLESHEET);

    if (!hasHeeboStylesheet) {
      const fontLink: HTMLLinkElement = document.createElement('link');
      fontLink.rel = 'stylesheet';
      fontLink.href = HEEBO_FONT_STYLESHEET;
      document.head.appendChild(fontLink);
    }
  }, []);

  React.useEffect(() => {
    if (isPaused || props.isLoading || props.loadError || window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
      return undefined;
    }

    const timer: number = window.setInterval(() => {
      setBannerState((current) => ({
        active: current.active === 0 ? 1 : 0,
        leaving: current.active
      }));
    }, 7000);

    return () => window.clearInterval(timer);
  }, [isPaused, props.isLoading, props.loadError]);

  React.useEffect(() => {
    const leavingBanner: number | null = bannerState.leaving;
    if (leavingBanner === null) {
      return undefined;
    }

    const timer: number = window.setTimeout(() => {
      setBannerState((current) => current.leaving === leavingBanner
        ? { ...current, leaving: null }
        : current);
    }, 650);

    return () => window.clearTimeout(timer);
  }, [bannerState.leaving]);

  if (props.isLoading) {
    return <div className={styles.dataState} role="status">טוען את נתוני מסע הקליטה...</div>;
  }

  if (props.loadError) {
    return (
      <div className={`${styles.dataState} ${styles.dataError}`} role="alert">
        <strong>לא ניתן לטעון את נתוני מסע הקליטה.</strong>
        <span>{props.loadError}</span>
      </div>
    );
  }

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
        clusterProgress={props.clusterProgress}
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
          <div className={styles.bannerMessages} aria-live="polite" aria-atomic="true">
            <div
              className={`${styles.bannerMessage} ${getBannerMessageClass(0)}`}
              aria-hidden={activeBanner !== 0}
            >
              <strong>{bannerTitle}</strong>
            </div>
            <div
              className={`${styles.bannerMessage} ${getBannerMessageClass(1)}`}
              aria-hidden={activeBanner !== 1}
            >
              <strong>יום {currentDay} מתוך {totalDays}</strong>
              <span className={styles.bannerProgress} aria-label={`${currentDay} מתוך ${totalDays} ימים`}>
                <i style={{ width: `${(currentDay / totalDays) * 100}%` }} />
              </span>
            </div>
          </div>
          <div className={styles.bannerDots} role="group" aria-label="בחירת הודעה">
            <button
              type="button"
              aria-label="מסר החזון"
              aria-pressed={activeBanner === 0}
              onFocus={() => setIsPaused(true)}
              onBlur={() => setIsPaused(false)}
              onClick={() => showBanner(0)}
            />
            <button
              type="button"
              aria-label="ההתקדמות שלך"
              aria-pressed={activeBanner === 1}
              onFocus={() => setIsPaused(true)}
              onBlur={() => setIsPaused(false)}
              onClick={() => showBanner(1)}
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
              <p>{welcomeParagraph1}</p>
              <p>{welcomeParagraph2.replace('{totalDays}', String(totalDays))}</p>
              <p className={styles.slogan}>{slogan}</p>
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
            <h2>{storyTitle}</h2>
            <ol className={styles.timeline}>
              {props.timelineSteps.map((step: IHeroTimelineStep) => (
                <li className={getTimelineClass(step.variant)} key={step.id}>
                  {step.variant === 'now' && (
                    <span className={styles.timelineAvatar} aria-hidden="true">
                      <span className={styles.timelineAvatarPhoto}>
                        <svg width="36" height="36" viewBox="0 0 64 64">
                          <rect width="64" height="64" fill="#FCE7B8" />
                          <path d="M20 30C19 16 25 11 32 11C39 11 45 16 44 30L46 50C41 49 39 45 39 40L25 40C25 45 23 49 18 50Z" fill="#4A2E1F" />
                          <path d="M11 66C11 52 20.5 46 32 46C43.5 46 53 52 53 66Z" fill="#232A5C" />
                          <path d="M27.5 37h9v9.5c0 2.2-2 3.5-4.5 3.5s-4.5-1.3-4.5-3.5z" fill="#E8B896" />
                          <path d="M27.5 41.5c2.8 1.6 6.2 1.6 9 0V44c-2.8 1.4-6.2 1.4-9 0z" fill="#000" opacity=".08" />
                          <ellipse cx="32" cy="29" rx="10.5" ry="12.5" fill="#E8B896" />
                          <path d="M21.5 28C21 18 26 14.5 32 14.5C38 14.5 43 18 42.5 28C40 22 35 19.5 30 21C27 22 24 24.5 21.5 28Z" fill="#4A2E1F" />
                          <circle cx="28" cy="29.5" r="1.35" fill="#2A2438" />
                          <circle cx="36" cy="29.5" r="1.35" fill="#2A2438" />
                          <path d="M28.6 34.6Q32 37.2 35.4 34.6" fill="none" stroke="#A85A48" strokeWidth="1.4" strokeLinecap="round" />
                        </svg>
                      </span>
                    </span>
                  )}
                  <strong>{step.heading}</strong>
                  <span>{step.description}</span>
                </li>
              ))}
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
