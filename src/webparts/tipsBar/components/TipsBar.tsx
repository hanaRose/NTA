import * as React from 'react';
import styles from './TipsBar.module.scss';
import { IOnboardingTip } from '../../../services/tips/IOnboardingTip';
import { ITipsBarProps } from './ITipsBarProps';

const LightbulbIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true">
    <path d="M9 18h6M10 21h4m-2-19a7 7 0 0 0-4 12.75c.5.35 1 1.05 1 1.75h6c0-.7.5-1.4 1-1.75A7 7 0 0 0 12 2Z" />
  </svg>
);

const ListIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true">
    <path d="M9 6h12M9 12h12M9 18h12" />
    <path d="M3 6h.01M3 12h.01M3 18h.01" />
  </svg>
);

const CloseIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true">
    <path d="m6 6 12 12M18 6 6 18" />
  </svg>
);

const ExternalLinkIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true">
    <path d="M14 4h6v6M20 4l-9 9" />
    <path d="M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5" />
  </svg>
);

const CheckIcon: React.FC = () => (
  <svg viewBox="0 0 24 24" aria-hidden="true">
    <path d="m5 12 4 4L19 6" />
  </svg>
);

function formatJourneyDay(day: number): string {
  return day === 0 ? 'טרום קליטה' : `יום ${day}`;
}

const TipsBar: React.FC<ITipsBarProps> = (props: ITipsBarProps) => {
  const [readTipIds, setReadTipIds] = React.useState<number[]>([]);
  const [isListOpen, setIsListOpen] = React.useState<boolean>(false);
  const [activeTip, setActiveTip] = React.useState<IOnboardingTip | undefined>();
  const tips: IOnboardingTip[] = props.tips;
  const featuredTip: IOnboardingTip | undefined = tips.length > 0 ? tips[0] : undefined;
  const isFeaturedRead: boolean = featuredTip !== undefined &&
    readTipIds.indexOf(featuredTip.id) >= 0;
  const listLabel: string = props.isLoading
    ? 'טוען טיפים'
    : (featuredTip
      ? (isFeaturedRead ? 'הטיפ האחרון' : 'טיפ חדש')
      : (props.loadError ? 'שגיאה בטעינת הטיפים' : 'אין טיפים חדשים'));

  React.useEffect(() => {
    if (!isListOpen && !activeTip) {
      return undefined;
    }

    const handleKeyDown = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        setIsListOpen(false);
        setActiveTip(undefined);
      }
    };

    document.addEventListener('keydown', handleKeyDown);
    return () => document.removeEventListener('keydown', handleKeyDown);
  }, [activeTip, isListOpen]);

  const openTip = (tip: IOnboardingTip): void => {
    setIsListOpen(false);
    setActiveTip(tip);
  };

  const markTipRead = (): void => {
    if (activeTip && readTipIds.indexOf(activeTip.id) < 0) {
      setReadTipIds(readTipIds.concat(activeTip.id));
    }
    setActiveTip(undefined);
  };

  return (
    <div className={`${styles.tipBar} ${isFeaturedRead ? styles.read : styles.new}`} dir="rtl">
      <span className={styles.tipIcon} aria-hidden="true"><LightbulbIcon /></span>
      <div className={styles.tipText}>
        <div className={styles.tipLabel}>
          <strong>{listLabel}</strong>
          {featuredTip && <span>{formatJourneyDay(featuredTip.day)} · {featuredTip.category}</span>}
          {isFeaturedRead && <span className={styles.readLabel}>נקרא</span>}
        </div>
        <div className={styles.tipTitle}>
          {featuredTip
            ? featuredTip.title
            : (props.isLoading
              ? 'טוען את הטיפים...'
              : (props.loadError ? 'לא ניתן לטעון את הטיפים' : 'אין כרגע טיפים שפורסמו.'))}
        </div>
        {props.loadError && (
          <div className={styles.tipError} role="alert">
            לא ניתן לטעון את הטיפים: {props.loadError}
          </div>
        )}
      </div>
      <div className={styles.actions}>
        <button
          type="button"
          className={`${styles.readButton} ${isFeaturedRead ? styles.readButtonRead : ''}`}
          onClick={() => featuredTip && openTip(featuredTip)}
          disabled={!featuredTip || props.isLoading}
        >
          {isFeaturedRead ? 'לקריאה שוב' : 'לקריאת הטיפ'}
        </button>
        <button
          type="button"
          className={styles.allTipsButton}
          onClick={() => setIsListOpen(true)}
        >
          <ListIcon />
          <span>כל הטיפים ({tips.length})</span>
        </button>
      </div>

      {isListOpen && (
        <div className={styles.backdrop} onClick={() => setIsListOpen(false)}>
          <section
            className={`${styles.dialog} ${styles.listDialog}`}
            role="dialog"
            aria-modal="true"
            aria-label="הטיפים שלי"
            onClick={(event: React.MouseEvent<HTMLElement>) => event.stopPropagation()}
          >
            <button
              type="button"
              className={styles.closeButton}
              onClick={() => setIsListOpen(false)}
              aria-label="סגירה"
            >
              <CloseIcon />
            </button>
            <div className={styles.dialogHeading}>
              <h2>הטיפים שלי</h2>
              <p>כל הטיפים שקיבלת במסע עד עכשיו.</p>
            </div>
            {props.isLoading ? (
              <div className={styles.dialogMessage}>טוען את הטיפים...</div>
            ) : props.loadError ? (
              <div className={styles.dialogError} role="alert">
                לא ניתן לטעון את הטיפים: {props.loadError}
              </div>
            ) : tips.length === 0 ? (
              <div className={styles.dialogMessage}>אין כרגע טיפים שפורסמו.</div>
            ) : (
              <ul className={styles.tipList}>
                {tips.map((tip: IOnboardingTip) => (
                  <li key={tip.id}>
                    <span className={styles.listIcon} aria-hidden="true"><LightbulbIcon /></span>
                    <span className={styles.listText}>
                      <strong>{tip.title}</strong>
                      <span>
                        {formatJourneyDay(tip.day)} · {tip.category} ·
                        {' '}{readTipIds.indexOf(tip.id) >= 0 ? 'נקרא' : 'עוד לא נקרא'}
                      </span>
                    </span>
                    <button type="button" className={styles.listReadButton} onClick={() => openTip(tip)}>
                      לקריאה
                    </button>
                  </li>
                ))}
              </ul>
            )}
          </section>
        </div>
      )}

      {activeTip && (
        <div className={styles.backdrop} onClick={() => setActiveTip(undefined)}>
          <section
            className={styles.dialog}
            role="dialog"
            aria-modal="true"
            aria-label={activeTip.title}
            onClick={(event: React.MouseEvent<HTMLElement>) => event.stopPropagation()}
          >
            <div className={styles.tipDialogHeader}>
              <button
                type="button"
                className={styles.closeButton}
                onClick={() => setActiveTip(undefined)}
                aria-label="סגירה"
              >
                <CloseIcon />
              </button>
              <div className={styles.dialogTipMeta}>
                <span className={styles.dialogTipIcon} aria-hidden="true"><LightbulbIcon /></span>
                <span>טיפ · {formatJourneyDay(activeTip.day)} · {activeTip.category}</span>
              </div>
            </div>
            <div className={styles.tipDialogBody}>
              <h2 className={styles.dialogTipTitle}>{activeTip.title}</h2>
              <p className={styles.dialogTipText}>{activeTip.text}</p>
              {activeTip.showQr ? (
                <div className={styles.tipQr}>
                  {activeTip.qrImageUrl ? (
                    <img
                      className={styles.tipQrImage}
                      src={activeTip.qrImageUrl}
                      alt="קוד QR לפתיחת הקישור"
                    />
                  ) : (
                    <div className={styles.tipQrMissing} role="alert">
                      לא הוגדרה תמונת QR לטיפ הזה.
                    </div>
                  )}
                  <div className={styles.tipQrText}>
                    <strong>סורקים את הקוד בטלפון</strong>
                    <span>הקוד פותח את הקישור ישירות בנייד.</span>
                    {activeTip.link && (
                      <a className={styles.tipLink} href={activeTip.link} target="_blank" rel="noopener noreferrer">
                        <ExternalLinkIcon />
                        {activeTip.linkLabel || 'לפתיחת הקישור'}
                      </a>
                    )}
                  </div>
                </div>
              ) : activeTip.link ? (
                <a className={styles.tipLink} href={activeTip.link} target="_blank" rel="noopener noreferrer">
                  <ExternalLinkIcon />
                  {activeTip.linkLabel || 'לפתיחת הקישור'}
                </a>
              ) : null}
              <div className={styles.dialogFooter}>
                <button type="button" className={styles.ackButton} onClick={markTipRead}>
                  <CheckIcon />
                  קיבלתי
                </button>
                <small>הטיפ יישאר בדף הבית עד הטיפ הבא.</small>
              </div>
            </div>
          </section>
        </div>
      )}
    </div>
  );
};

export default TipsBar;
