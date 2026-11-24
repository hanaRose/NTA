import { IOnboardingTaskMetrics } from './IOnboardingTaskMetrics';

export interface IOnboardingTaskMetricsService {
  getForJourney(journeyId: number, currentDay: number): Promise<IOnboardingTaskMetrics>;
}
