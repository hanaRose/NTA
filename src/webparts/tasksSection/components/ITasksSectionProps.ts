import { IOnboardingTask } from '../../../services/onboarding/IOnboardingTask';

export interface ITasksSectionProps {
  currentDay: number;
  tasks: IOnboardingTask[];
  isLoading: boolean;
  loadError?: string;
  onCompleteTask: (task: IOnboardingTask) => Promise<number>;
  onUndoCompletion: (completionId: number) => Promise<void>;
}
