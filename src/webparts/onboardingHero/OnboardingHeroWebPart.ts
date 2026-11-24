import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import OnboardingHero from './components/OnboardingHero';
import { IOnboardingHeroProps } from './components/IOnboardingHeroProps';

export interface IOnboardingHeroWebPartProps {
  userName: string;
  userRole: string;
  brandTitle: string;
  brandSubtitle: string;
  currentDay: number;
  totalDays: number;
  completedPercent?: number;
  completedTasks: number;
  totalTasks: number;
  openTaskCount: number;
  learningMinutes: number;
  bannerTitle: string;
  mapUrl: string;
}

export default class OnboardingHeroWebPart extends BaseClientSideWebPart<IOnboardingHeroWebPartProps> {
  public render(): void {
    const totalTasks: number = Math.max(1, Number(this.properties.totalTasks) || 40);
    const storedCompletedTasks: number = Number(this.properties.completedTasks);
    const legacyPercent: number = Number(this.properties.completedPercent);
    const completedTasks: number = Number.isFinite(storedCompletedTasks)
      ? storedCompletedTasks
      : Math.round((Number.isFinite(legacyPercent) ? legacyPercent : 57) * totalTasks / 100);

    const element: React.ReactElement<IOnboardingHeroProps> = React.createElement(
      OnboardingHero,
      {
        userName: this.properties.userName || this.context.pageContext.user.displayName,
        userRole: this.properties.userRole,
        brandTitle: this.properties.brandTitle,
        brandSubtitle: this.properties.brandSubtitle,
        currentDay: this.properties.currentDay,
        totalDays: this.properties.totalDays,
        completedTasks,
        totalTasks,
        openTaskCount: this.properties.openTaskCount,
        learningMinutes: this.properties.learningMinutes,
        bannerTitle: this.properties.bannerTitle,
        mapUrl: this.properties.mapUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'עריכת התוכן וההתקדמות באזור הפתיחה'
          },
          groups: [
            {
              groupName: 'תוכן',
              groupFields: [
                PropertyPaneTextField('userName', {
                  label: 'שם לתצוגה (ריק = המשתמש המחובר)'
                }),
                PropertyPaneTextField('userRole', {
                  label: 'תפקיד'
                }),
                PropertyPaneTextField('brandTitle', {
                  label: 'כותרת המותג'
                }),
                PropertyPaneTextField('brandSubtitle', {
                  label: 'כותרת משנה'
                }),
                PropertyPaneTextField('bannerTitle', {
                  label: 'כותרת הבאנר'
                }),
                PropertyPaneTextField('mapUrl', {
                  label: 'קישור למפת המסע'
                })
              ]
            },
            {
              groupName: 'התקדמות',
              groupFields: [
                PropertyPaneTextField('currentDay', {
                  label: 'יום נוכחי'
                }),
                PropertyPaneTextField('totalDays', {
                  label: 'מספר ימים כולל'
                }),
                PropertyPaneTextField('completedTasks', {
                  label: 'משימות שהושלמו'
                }),
                PropertyPaneTextField('totalTasks', {
                  label: 'מספר משימות כולל'
                }),
                PropertyPaneTextField('openTaskCount', {
                  label: 'משימות פתוחות'
                }),
                PropertyPaneTextField('learningMinutes', {
                  label: 'דקות למידה פתוחות'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
