import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import JourneyMap from './components/JourneyMap';
import { IJourneyMapProps } from './components/IJourneyMapProps';

export interface IJourneyMapWebPartProps {
  currentDay: number;
}

export default class JourneyMapWebPart extends BaseClientSideWebPart<IJourneyMapWebPartProps> {
  public render(): void {
    const element: React.ReactElement<IJourneyMapProps> = React.createElement(
      JourneyMap,
      { currentDay: this.properties.currentDay }
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
            description: 'הגדרת היום הנוכחי במסע הקליטה'
          },
          groups: [
            {
              groupName: 'התקדמות במסע',
              groupFields: [
                PropertyPaneTextField('currentDay', {
                  label: 'היום הנוכחי במסע (1–100)'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
