import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import TipsBar from './components/TipsBar';
import { ITipsBarProps } from './components/ITipsBarProps';

export interface ITipsBarWebPartProps {
  tipTitle: string;
  tipText: string;
  tipDay: number;
  category: string;
  tipLink: string;
  tipLinkLabel: string;
}

export default class TipsBarWebPart extends BaseClientSideWebPart<ITipsBarWebPartProps> {
  public render(): void {
    const element: React.ReactElement<ITipsBarProps> = React.createElement(
      TipsBar,
      {
        tipTitle: this.properties.tipTitle,
        tipText: this.properties.tipText,
        tipDay: this.properties.tipDay,
        category: this.properties.category,
        tipLink: this.properties.tipLink,
        tipLinkLabel: this.properties.tipLinkLabel
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
            description: 'עריכת הטיפ שמוצג בסרגל'
          },
          groups: [
            {
              groupName: 'תוכן הטיפ',
              groupFields: [
                PropertyPaneTextField('tipTitle', {
                  label: 'כותרת'
                }),
                PropertyPaneTextField('tipText', {
                  label: 'תוכן',
                  multiline: true,
                  rows: 4
                }),
                PropertyPaneTextField('tipDay', {
                  label: 'יום במסע'
                }),
                PropertyPaneTextField('category', {
                  label: 'נושא'
                }),
                PropertyPaneTextField('tipLink', {
                  label: 'קישור לטיפ (לא חובה)'
                }),
                PropertyPaneTextField('tipLinkLabel', {
                  label: 'טקסט הקישור'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
