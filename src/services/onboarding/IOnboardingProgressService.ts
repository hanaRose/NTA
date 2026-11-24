import { IOnboardingProgress } from './IOnboardingProgress';

export interface IOnboardingProgressService {
  getForUser(userEmail: string): Promise<IOnboardingProgress>;
}
