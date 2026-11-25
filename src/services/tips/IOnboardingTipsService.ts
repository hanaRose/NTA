import { IOnboardingTip } from './IOnboardingTip';

export interface IOnboardingTipsService {
  getPublishedTips(): Promise<IOnboardingTip[]>;
}
