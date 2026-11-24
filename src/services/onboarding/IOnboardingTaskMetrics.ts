export interface IOnboardingClusterMetrics {
  name: string;
  totalTasks: number;
  completedTasks: number;
}

export interface IOnboardingTaskMetrics {
  completedTasks: number;
  totalTasks: number;
  openTaskCount: number;
  learningMinutes: number;
  clusterProgress: IOnboardingClusterMetrics[];
}

export const ONBOARDING_CLUSTER_NAMES: readonly string[] = [
  'תחנת מוצא',
  'מרכז בקרה',
  'תנאי הנסיעה',
  'חוקי הדרך',
  'מרכז לוגיסטי',
  'המסלול האישי'
];
