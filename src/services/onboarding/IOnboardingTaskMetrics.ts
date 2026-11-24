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

export const ONBOARDING_CLUSTER_COLORS: Readonly<Record<string, string>> = {
  'תחנת מוצא': '#e84a3c',
  'מרכז בקרה': '#9b4dca',
  'תנאי הנסיעה': '#4fa83d',
  'חוקי הדרך': '#f2a93b',
  'מרכז לוגיסטי': '#1fb5a6',
  'המסלול האישי': '#3d9be9'
};
