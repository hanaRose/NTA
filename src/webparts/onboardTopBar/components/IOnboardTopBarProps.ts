import { IOnboardingClusterMetrics } from '../../../services/onboarding/IOnboardingTaskMetrics';

export interface IOnboardTopBarProps {
  brandTitle: string;
  brandSubtitle: string;
  displayName: string;
  userRole: string;
  currentDay: number;
  totalDays: number;
  completedTasks: number;
  totalTasks: number;
  clusterProgress?: IOnboardingClusterMetrics[];
  standalone?: boolean;
}
