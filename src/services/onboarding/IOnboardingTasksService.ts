import { IOnboardingTask } from './IOnboardingTask';

export interface IOnboardingTasksService {
  getIncompleteForJourney(journeyId: number): Promise<IOnboardingTask[]>;
  createCompletion(
    journeyId: number,
    taskId: number,
    taskTitle: string,
    journeyDay: number
  ): Promise<number>;
  deleteCompletion(completionId: number): Promise<void>;
}
