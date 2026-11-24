import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import TasksSection from './components/TasksSection';
import { ITasksSectionProps } from './components/ITasksSectionProps';

export interface ITasksSectionWebPartProps {
  currentDay: number;
  payrollUrl: string;
  cyberUrl: string;
  safetyGearUrl: string;
  sapUrl: string;
  biUrl: string;
  knowledgeUrl: string;
}

export default class TasksSectionWebPart extends BaseClientSideWebPart<ITasksSectionWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ITasksSectionProps> = React.createElement(
      TasksSection,
      {
        currentDay: this.properties.currentDay,
        payrollUrl: this.properties.payrollUrl,
        cyberUrl: this.properties.cyberUrl,
        safetyGearUrl: this.properties.safetyGearUrl,
        sapUrl: this.properties.sapUrl,
        biUrl: this.properties.biUrl,
        knowledgeUrl: this.properties.knowledgeUrl
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
            description: 'הגדרת היום במסע וקישורי המשימות'
          },
          groups: [
            {
              groupName: 'התקדמות',
              groupFields: [
                PropertyPaneTextField('currentDay', {
                  label: 'היום הנוכחי במסע'
                })
              ]
            },
            {
              groupName: 'קישורים למשימות פתוחות',
              groupFields: [
                PropertyPaneTextField('payrollUrl', {
                  label: 'נהלי שכר ונוכחות'
                }),
                PropertyPaneTextField('cyberUrl', {
                  label: 'לומדת אבטחת מידע וסייבר'
                }),
                PropertyPaneTextField('safetyGearUrl', {
                  label: 'קבלת ציוד בטיחות'
                })
              ]
            },
            {
              groupName: 'קישורים למשימות הבאות',
              groupFields: [
                PropertyPaneTextField('sapUrl', {
                  label: 'הדרכת SAP'
                }),
                PropertyPaneTextField('biUrl', {
                  label: 'הדרכת BI'
                }),
                PropertyPaneTextField('knowledgeUrl', {
                  label: 'מערך ניהול הידע בארגון'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
