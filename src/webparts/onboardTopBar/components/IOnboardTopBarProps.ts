export interface IOnboardTopBarProps {
  brandTitle: string;
  brandSubtitle: string;
  displayName: string;
  userRole: string;
  currentDay: number;
  totalDays: number;
  completedTasks: number;
  totalTasks: number;
  standalone?: boolean;
}
