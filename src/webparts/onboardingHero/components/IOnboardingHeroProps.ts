import { IOnboardingClusterMetrics } from '../../../services/onboarding/IOnboardingTaskMetrics';

export type HeroTimelineStepVariant = 'past' | 'red' | 'now' | 'next';

export interface IHeroTimelineStep {
  id: string;
  heading: string;
  description: string;
  variant: HeroTimelineStepVariant;
}

export interface IOnboardingHeroProps {
  userName: string;
  userRole: string;
  brandTitle: string;
  brandSubtitle: string;
  currentDay: number;
  totalDays: number;
  isLoading: boolean;
  loadError?: string;
  completedTasks: number;
  totalTasks: number;
  clusterProgress: IOnboardingClusterMetrics[];
  openTaskCount: number;
  learningMinutes: number;
  bannerTitle: string;
  mapUrl: string;
  welcomeParagraph1: string;
  welcomeParagraph2: string;
  slogan: string;
  storyTitle: string;
  timelineSteps: IHeroTimelineStep[];
}
