export interface IOnboardingTask {
  id: number;
  title: string;
  taskContentType: string;
  journeyDay: number;
  duration?: number;
  clusterName: string;
  createdAt: Date;
  url: string;
}
