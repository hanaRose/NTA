import { IOnboardingTip } from '../../../services/tips/IOnboardingTip';

export interface ITipsBarProps {
  tips: IOnboardingTip[];
  isLoading: boolean;
  loadError?: string;
}
