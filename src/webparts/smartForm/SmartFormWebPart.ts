import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  PropertyPaneTextField,
  PropertyPaneLabel
} from '@microsoft/sp-webpart-base';

import FormApp, { FormAppProps } from './components/FormApp';

export interface ISmartFormWebPartProps {
  /** שם רשימת ה-PMO לעריכה */
  pmoListTitle: string;
  /** שם עמודת ה-Lookup ב-PMO שמקשרת ל-Integration */
  pmoIntegrationLookupName?: string;
  /** JSON של שלבים -> שמות שדות (Display/Internal) */
  stepsConfigJson?: string;
}

export default class SmartFormWebPart extends BaseClientSideWebPart<ISmartFormWebPartProps> {

  public render(): void {
    // פרש את stepsConfigJson אם קיים
    let stepsConfig: Record<string, string[]> = {};
    const raw = this.properties.stepsConfigJson;
    if (raw && raw.trim().length > 0) {
      try {
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === 'object') {
          stepsConfig = parsed as Record<string, string[]>;
        }
      } catch (e) {
        console.warn('[SmartForm] stepsConfigJson parse failed:', e);
      }
    }

    const props: FormAppProps = {
      context: this.context,
      pmoListTitle: this.properties.pmoListTitle || 'PMO decisions',
      pmoIntegrationLookupName: this.properties.pmoIntegrationLookupName || 'Integration',
      stepsConfig
    };

    const element = React.createElement(FormApp, props);
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration() {
    return {
      pages: [
        {
          header: { description: 'SmartForm – הגדרות' },
          groups: [
            {
              groupName: 'רשימות וקישור',
              groupFields: [
                PropertyPaneTextField('pmoListTitle', {
                  label: 'שם רשימת PMO (עריכה)',
                  placeholder: 'PMO decisions'
                }),
                PropertyPaneTextField('pmoIntegrationLookupName', {
                  label: 'שם עמודת ה-Lookup ב-PMO שמקשרת ל-Integration',
                  placeholder: 'Integration'
                })
              ]
            },
            {
              groupName: 'שלבים (Steps)',
              groupFields: [
                PropertyPaneLabel('', { text: 'Steps Config (JSON): {"שלב":[ "שם שדה", "שם שדה" ], ...}' }),
                PropertyPaneTextField('stepsConfigJson', {
                  label: 'Steps Config (JSON)',
                  multiline: true,
                  resizable: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
