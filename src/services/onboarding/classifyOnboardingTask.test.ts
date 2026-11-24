import { IOnboardingTask } from './IOnboardingTask';
import { classifyOnboardingTask } from './classifyOnboardingTask';

const today: Date = new Date(2026, 10, 24, 12);

function createTask(journeyDay: number, createdAt: Date): IOnboardingTask {
  return {
    id: 1,
    title: 'Task',
    taskContentType: 'סרטון',
    url: 'https://contoso.com/task',
    journeyDay,
    clusterName: 'תחנת מוצא',
    createdAt
  };
}

describe('classifyOnboardingTask', () => {
  it('marks an earlier due task created today as new today', () => {
    expect(classifyOnboardingTask(
      createTask(4, new Date(2026, 10, 24, 8)),
      5,
      today
    )).toBe('newToday');
  });

  it('marks an earlier due task created on a previous date as overdue', () => {
    expect(classifyOnboardingTask(
      createTask(4, new Date(2026, 10, 23, 23, 59)),
      5,
      today
    )).toBe('overdue');
  });

  it('leaves a task due today untagged and keeps future tasks upcoming', () => {
    expect(classifyOnboardingTask(
      createTask(5, new Date(2026, 10, 24, 8)),
      5,
      today
    )).toBe('dueToday');
    expect(classifyOnboardingTask(
      createTask(6, new Date(2026, 10, 24, 8)),
      5,
      today
    )).toBe('upcoming');
  });
});
