import { IOnboardingTask } from './IOnboardingTask';

export type OnboardingTaskStatus = 'overdue' | 'newToday' | 'dueToday' | 'upcoming';

function isSameLocalDay(firstDate: Date, secondDate: Date): boolean {
  return firstDate.getFullYear() === secondDate.getFullYear() &&
    firstDate.getMonth() === secondDate.getMonth() &&
    firstDate.getDate() === secondDate.getDate();
}

export function classifyOnboardingTask(
  task: IOnboardingTask,
  currentDay: number,
  today: Date
): OnboardingTaskStatus {
  if (task.journeyDay > currentDay) {
    return 'upcoming';
  }
  if (task.journeyDay === currentDay) {
    return 'dueToday';
  }
  return isSameLocalDay(task.createdAt, today) ? 'newToday' : 'overdue';
}
