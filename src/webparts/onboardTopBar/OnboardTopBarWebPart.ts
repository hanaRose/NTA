import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import OnboardTopBar from './components/OnboardTopBar';
import { IOnboardTopBarProps } from './components/IOnboardTopBarProps';

export interface IOnboardTopBarWebPartProps {
  brandTitle: string;
  brandSubtitle: string;
  displayName: string;
  userRole: string;
  currentDay: number;
  totalDays: number;
  completedTasks: number;
  totalTasks: number;
}

export default class OnboardTopBarWebPart extends BaseClientSideWebPart<IOnboardTopBarWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IOnboardTopBarProps> = React.createElement(
      OnboardTopBar,
      {
        brandTitle: this.properties.brandTitle,
        brandSubtitle: this.properties.brandSubtitle,
        displayName: this.properties.displayName || this.context.pageContext.user.displayName,
        userRole: this.properties.userRole,
        currentDay: this.properties.currentDay,
        totalDays: this.properties.totalDays,
        completedTasks: this.properties.completedTasks,
        totalTasks: this.properties.totalTasks,
        standalone: true
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
            description: 'עריכת התוכן שמוצג בסרגל העליון'
          },
          groups: [
            {
              groupName: 'פרטי המותג והמשתמש',
              groupFields: [
                PropertyPaneTextField('brandTitle', {
                  label: 'כותרת'
                }),
                PropertyPaneTextField('brandSubtitle', {
                  label: 'כותרת משנה'
                }),
                PropertyPaneTextField('displayName', {
                  label: 'שם העובד/ת (ריק = המשתמש המחובר)'
                }),
                PropertyPaneTextField('userRole', {
                  label: 'תפקיד'
                })
              ]
            },
            {
              groupName: 'התקדמות במסע',
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
