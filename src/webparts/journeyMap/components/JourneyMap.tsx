import * as React from 'react';
import styles from './JourneyMap.module.scss';
import { IJourneyMapProps } from './IJourneyMapProps';

type MilestoneStatus = 'done' | 'open' | 'late' | 'locked';

interface IMilestone {
  id: string;
  day: number;
  late?: boolean;
}

interface IStation {
  id: string;
  name: string;
  description: string;
  color: string;
  ink: string;
  milestones: IMilestone[];
}

interface IMilestoneGroup {
  x: number;
  milestones: IMilestone[];
}

const MAP_WIDTH: number = 1240;
const AXIS_RIGHT: number = 985;
const AXIS_LEFT: number = 50;
const FIRST_ROW_Y: number = 112;
const ROW_GAP: number = 62;
const MAP_HEIGHT: number = FIRST_ROW_Y + 5 * ROW_GAP + 44;
const stations: IStation[] = [
  {
    id: 'origin',
    name: 'תחנת מוצא',
    description: 'שער הכניסה',
    color: '#E84A3C',
    ink: '#FFFFFF',
    milestones: [
      { id: 'o1', day: 0 },
      { id: 'o2', day: 0 },
      { id: 'o3', day: 0 },
      { id: 'o4', day: 1 },
      { id: 'o5', day: 45 }
    ]
  },
  {
    id: 'control',
    name: 'מרכז בקרה',
    description: 'מערכות ותהליכי עבודה',
    color: '#9B4DCA',
    ink: '#FFFFFF',
    milestones: [
      { id: 's1', day: 1 },
      { id: 's2', day: 1 },
      { id: 's3', day: 1 },
      { id: 's4', day: 2 },
      { id: 's5', day: 7 },
      { id: 's6', day: 10 },
      { id: 's7', day: 13 },
      { id: 's8', day: 16 },
      { id: 's9', day: 19 },
      { id: 's10', day: 21 },
      { id: 's11', day: 22 },
      { id: 's12', day: 25 }
    ]
  },
  {
    id: 'terms',
    name: 'תנאי הנסיעה',
    description: 'שכר, נוכחות וזכויות',
    color: '#4FA83D',
    ink: '#FFFFFF',
    milestones: [
      { id: 'p1', day: 0 },
      { id: 'p2', day: 1 },
      { id: 'p3', day: 3 },
      { id: 'p4', day: 7, late: true },
      { id: 'p5', day: 29 },
      { id: 'p6', day: 29 },
      { id: 'p7', day: 30 },
      { id: 'p8', day: 60 }
    ]
  },
  {
    id: 'route',
    name: 'חוקי הדרך',
    description: 'ציות ורגולציה',
    color: '#F2A93B',
    ink: '#232A5C',
    milestones: [
      { id: 'c1', day: 1 },
      { id: 'c2', day: 2 },
      { id: 'c3', day: 5 },
      { id: 'c4', day: 5 },
      { id: 'c5', day: 10, late: true },
      { id: 'c6', day: 14 },
      { id: 'c7', day: 35 }
    ]
  },
  {
    id: 'logistics',
    name: 'מרכז לוגיסטי',
    description: 'ציוד וסביבת עבודה',
    color: '#1FB5A6',
    ink: '#FFFFFF',
    milestones: [
      { id: 'e1', day: 0 },
      { id: 'e2', day: 0 },
      { id: 'e3', day: 0 },
      { id: 'e4', day: 1 },
      { id: 'e5', day: 2 }
    ]
  },
  {
    id: 'personal',
    name: 'המסלול האישי',
    description: 'פיתוח אישי',
    color: '#3D9BE9',
    ink: '#FFFFFF',
    milestones: [
      { id: 'd1', day: 35 },
      { id: 'd2', day: 42 },
      { id: 'd3', day: 100 }
    ]
  }
];

function readDay(value: number): number {
  const parsedValue: number = Math.floor(Number(value));
  return Number.isFinite(parsedValue)
    ? Math.max(0, Math.min(100, parsedValue))
    : 14;
}

function getX(day: number): number {
  return AXIS_RIGHT - (AXIS_RIGHT - AXIS_LEFT) * Math.pow(Math.max(0, day) / 100, 0.6);
}

function getStatus(milestone: IMilestone, currentDay: number): MilestoneStatus {
  if (milestone.day > currentDay) {
    return 'locked';
  }
  if (milestone.day === currentDay) {
    return 'open';
  }
  return milestone.late ? 'late' : 'done';
}

function getGroups(station: IStation): IMilestoneGroup[] {
  const groups: IMilestoneGroup[] = [];

  station.milestones.slice().sort((a: IMilestone, b: IMilestone) => a.day - b.day).forEach(
    (milestone: IMilestone) => {
      const x: number = getX(milestone.day);
      const previousGroup: IMilestoneGroup | undefined = groups[groups.length - 1];
      if (previousGroup && Math.abs(previousGroup.x - x) < 16) {
        previousGroup.milestones.push(milestone);
      } else {
        groups.push({ x, milestones: [milestone] });
      }
    }
  );

  return groups;
}

const Legend: React.FC = () => (
  <div className={styles.legend} aria-label="מקרא">
    <span>
      <svg viewBox="0 0 16 16" aria-hidden="true">
        <circle cx="8" cy="8" r="6.5" fill="#232A5C" />
        <path d="m5.2 8.2 1.9 1.9 3.8-4" />
      </svg>
      הסתיימה
    </span>
    <span>
      <svg viewBox="0 0 16 16" aria-hidden="true">
        <circle cx="8" cy="8" r="5.5" />
      </svg>
      פתוחה היום
    </span>
    <span>
      <svg viewBox="0 0 16 16" aria-hidden="true">
        <circle cx="8" cy="8" r="6.5" fill="#E8414F" />
        <path d="M8 4.8v3.8" />
        <circle cx="8" cy="11.1" r="1" fill="#fff" stroke="none" />
      </svg>
      באיחור
    </span>
    <span>
      <svg viewBox="0 0 34 12" aria-hidden="true">
        <path d="M2 6h22" />
        <circle cx="28.5" cy="6" r="4" />
      </svg>
      נעולה, תיפתח במועדה
    </span>
    <span>
      <svg viewBox="0 0 16 16" aria-hidden="true">
        <circle cx="8" cy="8" r="7" fill="#232A5C" />
        <text x="8" y="11.3" textAnchor="middle">2</text>
      </svg>
      כמה משימות באותה תחנה
    </span>
  </div>
);

const TrainIcon: React.FC = () => (
  <svg className={styles.trainIcon} viewBox="-32 -31 64 35" aria-hidden="true">
    <path d="M-20 -23H24Q30 -23 30 -17V-5Q30 -1 26 -1H-27Q-31 -1 -30.4 -4.5L-27 -17Q-25.5 -23 -20 -23Z" fill="#fff" />
    <path d="M-21 -19h8v7h-8zM-10 -19h8v7h-8zM1 -19h8v7h-8zM12 -19h8v7h-8zM-27.5 -10.5L-25 -18.5H-23V-10.5Z" fill="#232A5C" />
    <rect x="-29.5" y="-8.5" width="59" height="3" fill="#F2B33D" />
    <path d="M-2 -23L-6 -30H6" fill="none" stroke="#fff" strokeWidth="2.4" strokeLinecap="round" />
  </svg>
);

const JourneyMap: React.FC<IJourneyMapProps> = (props: IJourneyMapProps) => {
  const currentDay: number = readDay(props.currentDay);
  const currentX: number = getX(currentDay);
  const trackBottom: number = FIRST_ROW_Y + 5 * ROW_GAP + 24;

  return (
    <section id="journeyMap" className={styles.journeyMap} dir="rtl" aria-label="מפת מסע הקליטה">
      <header className={styles.header}>
        <div className={styles.heading}>
          <span>הסיפור שלך · 100 ימי קליטה</span>
          <h2>מפת המסע</h2>
        </div>
        <Legend />
      </header>

      <div className={styles.mapCard}>
        <svg className={styles.cornerDecoration} viewBox="0 0 150 110" aria-hidden="true">
          <path d="M0 0H64L0 64Z" fill="#232A5C" />
          <path d="M78 0H92L0 92V78Z" fill="#A9C8EE" />
          <path d="M104 0H112L0 112V104Z" fill="#F2B33D" />
        </svg>

        <div className={styles.mapScroll}>
          <div className={styles.mapHost} style={{ width: MAP_WIDTH, height: MAP_HEIGHT }}>
            <svg
              className={styles.mapSvg}
              width={MAP_WIDTH}
              height={MAP_HEIGHT}
              viewBox={`0 0 ${MAP_WIDTH} ${MAP_HEIGHT}`}
              role="img"
              aria-labelledby="journey-map-title journey-map-description"
            >
              <title id="journey-map-title">מפת מסע הקליטה</title>
              <desc id="journey-map-description">
                שישה מסלולים לאורך מאה ימי קליטה. היום הנוכחי הוא יום {currentDay}.
              </desc>
              <path d={`M${AXIS_RIGHT} 48H${AXIS_LEFT}`} className={styles.axisLine} />
              <path d={`M${AXIS_RIGHT} 48H${currentX}`} className={styles.axisProgress} />
              <circle cx={AXIS_RIGHT} cy="48" r="6.5" className={styles.axisEnd} />
              <circle cx={AXIS_LEFT} cy="48" r="6.5" className={styles.axisStart} />
              <path
                d={`M${currentX} 62V${trackBottom}`}
                className={styles.currentGuide}
              />

              {stations.map((station: IStation, index: number) => {
                const y: number = FIRST_ROW_Y + index * ROW_GAP;
                const groups: IMilestoneGroup[] = getGroups(station);
                const firstX: number = groups.length ? groups[0].x : 0;
                const lastX: number = groups.length ? groups[groups.length - 1].x : 0;

                return (
                  <g key={station.id}>
                    {firstX > currentX && (
                      <path
                        d={`M${firstX} ${y}H${Math.max(lastX, currentX)}`}
                        className={styles.activeTrack}
                        style={{ stroke: station.color }}
                      />
                    )}
                    {lastX < currentX && (
                      <path
                        d={`M${Math.min(firstX, currentX)} ${y}H${lastX}`}
                        className={styles.futureTrack}
                        style={{ stroke: station.color }}
                      />
                    )}

                    {groups.map((group: IMilestoneGroup, groupIndex: number) => {
                      const statuses: MilestoneStatus[] = group.milestones.map(
                        (milestone: IMilestone) => getStatus(milestone, currentDay)
                      );
                      const status: MilestoneStatus = statuses.indexOf('late') >= 0
                        ? 'late'
                        : (statuses.indexOf('open') >= 0
                          ? 'open'
                          : (statuses.every((value: MilestoneStatus) => value === 'done') ? 'done' : 'locked'));
                      const label: string = group.milestones.length === 1
                        ? `יום ${group.milestones[0].day}`
                        : `${group.milestones.length} משימות בתחנה`;

                      return (
                        <g key={`${station.id}-${groupIndex}`} aria-label={`${station.name}, ${label}`}>
                          <title>{`${station.name} · ${label}`}</title>
                          {status === 'done' && (
                            <React.Fragment>
                              <circle cx={group.x} cy={y} r="7.5" className={styles.doneNode} style={{ fill: station.color }} />
                              <path
                                d={`M${group.x - 3.2} ${y + 0.2}l2.1 2.1 4.2-4.4`}
                                className={styles.checkMark}
                                style={{ stroke: station.ink }}
                              />
                            </React.Fragment>
                          )}
                          {status === 'open' && (
                            <React.Fragment>
                              <circle cx={group.x} cy={y} r="15" className={styles.openHalo} style={{ fill: station.color }} />
                              <circle cx={group.x} cy={y} r="8" className={styles.openNode} style={{ stroke: station.color }} />
                            </React.Fragment>
                          )}
                          {status === 'late' && (
                            <React.Fragment>
                              <circle cx={group.x} cy={y} r="13" className={styles.lateHalo} />
                              <circle cx={group.x} cy={y} r="8.5" className={styles.lateNode} />
                              <path d={`M${group.x} ${y - 3.8}V${y + 0.8}`} className={styles.lateMark} />
                              <circle cx={group.x} cy={y + 3.8} r="1.1" className={styles.lateDot} />
                            </React.Fragment>
                          )}
                          {status === 'locked' && (
                            <circle
                              cx={group.x}
                              cy={y}
                              r="6"
                              className={styles.lockedNode}
                              style={{ stroke: station.color }}
                            />
                          )}
                          {group.milestones.length > 1 && (
                            <g className={styles.groupBadge}>
                              <circle cx={group.x + 8} cy={y - 9} r="7" />
                              <text x={group.x + 8} y={y - 5.6} textAnchor="middle">
                                {group.milestones.length}
                              </text>
                            </g>
                          )}
                        </g>
                      );
                    })}
                  </g>
                );
              })}
            </svg>

            <div className={styles.axisLabel} style={{ right: MAP_WIDTH - AXIS_RIGHT }}>טרום קליטה</div>
            <div className={styles.axisLabel} style={{ right: MAP_WIDTH - AXIS_LEFT }}>יום 100</div>
            <div className={styles.todayTag} style={{ right: MAP_WIDTH - currentX }}>
              <TrainIcon />
              <span>את כאן · יום {currentDay}</span>
            </div>

            {stations.map((station: IStation, index: number) => {
              const y: number = FIRST_ROW_Y + index * ROW_GAP;
              const completedCount: number = station.milestones.filter(
                (milestone: IMilestone) => getStatus(milestone, currentDay) === 'done'
              ).length;
              const progress: number = station.milestones.length
                ? Math.round((completedCount / station.milestones.length) * 100)
                : 0;

              return (
                <div className={styles.lane} style={{ top: y - 21 }} key={station.id}>
                  <div className={styles.laneHeading}>
                    <svg viewBox="0 0 22 10" aria-hidden="true">
                      <path d="M2 5H14" style={{ stroke: station.color }} />
                      <circle cx="17" cy="5" r="3.4" style={{ stroke: station.color }} />
                    </svg>
                    <b>{station.name}</b>
                    <em>{completedCount}/{station.milestones.length}</em>
                  </div>
                  <div className={styles.laneDescription}>{station.description}</div>
                  <div className={styles.laneProgress}>
                    <span style={{ width: `${progress}%`, backgroundColor: station.color }} />
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </section>
  );
};

export default JourneyMap;
