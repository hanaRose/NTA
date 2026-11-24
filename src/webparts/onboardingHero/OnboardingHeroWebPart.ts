import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { IOnboardingProgress } from '../../services/onboarding/IOnboardingProgress';
import { IOnboardingProgressService } from '../../services/onboarding/IOnboardingProgressService';
import {
  IOnboardingTaskMetrics,
  ONBOARDING_CLUSTER_NAMES
} from '../../services/onboarding/IOnboardingTaskMetrics';
import { IOnboardingTaskMetricsService } from '../../services/onboarding/IOnboardingTaskMetricsService';
import SharePointOnboardingProgressService from '../../services/onboarding/SharePointOnboardingProgressService';
import SharePointOnboardingTaskMetricsService from '../../services/onboarding/SharePointOnboardingTaskMetricsService';
import OnboardingHero from './components/OnboardingHero';
import { IHeroTimelineStep, IOnboardingHeroProps } from './components/IOnboardingHeroProps';

export interface IOnboardingHeroWebPartProps {
  userRole: string;
  brandTitle: string;
  brandSubtitle: string;
  totalDays: number;
  bannerTitle: string;
  mapUrl: string;
  welcomeParagraph1: string;
  welcomeParagraph2: string;
  slogan: string;
  storyTitle: string;
  step1Heading: string;
  step1Description: string;
  step2Heading: string;
  step2Description: string;
  step3Heading: string;
  step3Description: string;
  step4Heading: string;
  step4Description: string;
}

export default class OnboardingHeroWebPart extends BaseClientSideWebPart<IOnboardingHeroWebPartProps> {
  private _progressService: IOnboardingProgressService | undefined;
  private _taskMetricsService: IOnboardingTaskMetricsService | undefined;
  private _progress: IOnboardingProgress | undefined;
  private _taskMetrics: IOnboardingTaskMetrics = {
    completedTasks: 0,
    totalTasks: 0,
    openTaskCount: 0,
    learningMinutes: 0,
    clusterProgress: ONBOARDING_CLUSTER_NAMES.map((name: string) => ({
      name,
      totalTasks: 0,
      completedTasks: 0
    }))
  };
  private _progressError: string | undefined;
  private _isProgressLoading: boolean = true;
  private _isHeroDisposed: boolean = false;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._progressService = new SharePointOnboardingProgressService(
      this.context.spHttpClient,
      this.context.pageContext.web.absoluteUrl
    );
    this._taskMetricsService = new SharePointOnboardingTaskMetricsService(
      this.context.spHttpClient,
      this.context.pageContext.web.absoluteUrl
    );
    this.loadProgress().catch((error: unknown) => {
      console.error('Unexpected error while loading onboarding progress.', error);
    });
  }

  public render(): void {
    const totalDays: number = Math.max(1, Number(this.properties.totalDays) || 100);
    const timelineSteps: IHeroTimelineStep[] = [
      {
        id: 'step-1',
        heading: this.properties.step1Heading || '1925',
        description: this.properties.step1Description || 'נולד החלום על רכבת קלה בתל אביב',
        variant: 'past'
      },
      {
        id: 'step-2',
        heading: this.properties.step2Heading || '2023',
        description: this.properties.step2Description || 'נפתח הקו האדום',
        variant: 'red'
      },
      {
        id: 'step-3',
        heading: this.properties.step3Heading || 'היום',
        description: this.properties.step3Description || 'הקווים הירוק והסגול בבנייה, ואת מצטרפת',
        variant: 'now'
      },
      {
        id: 'step-4',
        heading: this.properties.step4Heading || 'הפרק הבא',
        description: this.properties.step4Description || 'המטרו',
        variant: 'next'
      }
    ];

    const element: React.ReactElement<IOnboardingHeroProps> = React.createElement(
      OnboardingHero,
      {
        userName: this.context.pageContext.user.displayName,
        userRole: this.properties.userRole,
        brandTitle: this.properties.brandTitle,
        brandSubtitle: this.properties.brandSubtitle,
        currentDay: this._progress ? this._progress.currentDay : 1,
        totalDays,
        isLoading: this._isProgressLoading,
        loadError: this._progressError,
        completedTasks: this._taskMetrics.completedTasks,
        totalTasks: this._taskMetrics.totalTasks,
        clusterProgress: this._taskMetrics.clusterProgress,
        openTaskCount: this._taskMetrics.openTaskCount,
        learningMinutes: this._taskMetrics.learningMinutes,
        bannerTitle: this.properties.bannerTitle,
        mapUrl: this.properties.mapUrl,
        welcomeParagraph1: this.properties.welcomeParagraph1,
        welcomeParagraph2: this.properties.welcomeParagraph2,
        slogan: this.properties.slogan,
        storyTitle: this.properties.storyTitle,
        timelineSteps
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    this._isHeroDisposed = true;
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private async loadProgress(): Promise<void> {
    try {
      if (!this._progressService) {
        throw new Error('The onboarding progress service is not available.');
      }
      this._progress = await this._progressService.getForUser(this.context.pageContext.user.email);
      if (!this._taskMetricsService) {
        throw new Error('The onboarding task metrics service is not available.');
      }
      this._taskMetrics = await this._taskMetricsService.getForJourney(
        this._progress.journeyId,
        this._progress.currentDay
      );
    } catch (error) {
      this._progressError = error instanceof Error ? error.message : String(error);
    } finally {
      this._isProgressLoading = false;
      if (!this._isHeroDisposed) {
        this.render();
      }
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'עריכת הטקסטים וסיפור המסגרת באזור הפתיחה'
          },
          groups: [
            {
              groupName: 'כותרות ותוכן',
              groupFields: [
                PropertyPaneTextField('brandTitle', { label: 'כותרת המותג' }),
                PropertyPaneTextField('brandSubtitle', { label: 'כותרת משנה' }),
                PropertyPaneTextField('userRole', { label: 'תפקיד לתצוגה' }),
                PropertyPaneTextField('bannerTitle', { label: 'כותרת הבאנר' }),
                PropertyPaneTextField('mapUrl', { label: 'קישור למפת המסע' }),
                PropertyPaneTextField('welcomeParagraph1', {
                  label: 'טקסט פתיחה',
                  multiline: true
                }),
                PropertyPaneTextField('welcomeParagraph2', {
                  label: 'טקסט המסע',
                  multiline: true
                }),
                PropertyPaneTextField('slogan', { label: 'משפט סיום' }),
                PropertyPaneTextField('storyTitle', { label: 'כותרת סיפור המסגרת' })
              ]
            },
            {
              groupName: 'שלבי סיפור המסגרת',
              groupFields: [
                PropertyPaneTextField('step1Heading', { label: 'שלב 1 — כותרת' }),
                PropertyPaneTextField('step1Description', { label: 'שלב 1 — תיאור' }),
                PropertyPaneTextField('step2Heading', { label: 'שלב 2 — כותרת' }),
                PropertyPaneTextField('step2Description', { label: 'שלב 2 — תיאור' }),
                PropertyPaneTextField('step3Heading', { label: 'שלב 3 — כותרת' }),
                PropertyPaneTextField('step3Description', { label: 'שלב 3 — תיאור' }),
                PropertyPaneTextField('step4Heading', { label: 'שלב 4 — כותרת' }),
                PropertyPaneTextField('step4Description', { label: 'שלב 4 — תיאור' })
              ]
            },
            {
              groupName: 'משך המסע',
              groupFields: [
                PropertyPaneTextField('totalDays', { label: 'מספר ימים כולל' })
              ]
            }
          ]
        }
      ]
    };
  }
}
