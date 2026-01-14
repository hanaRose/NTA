var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart, PropertyPaneTextField, PropertyPaneLabel } from '@microsoft/sp-webpart-base';
import FormApp from './components/FormApp';
var SmartFormWebPart = /** @class */ (function (_super) {
    __extends(SmartFormWebPart, _super);
    function SmartFormWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    SmartFormWebPart.prototype.render = function () {
        // פרש את stepsConfigJson אם קיים
        var stepsConfig = {};
        var raw = this.properties.stepsConfigJson;
        if (raw && raw.trim().length > 0) {
            try {
                var parsed = JSON.parse(raw);
                if (parsed && typeof parsed === 'object') {
                    stepsConfig = parsed;
                }
            }
            catch (e) {
                console.warn('[SmartForm] stepsConfigJson parse failed:', e);
            }
        }
        var props = {
            context: this.context,
            pmoListTitle: this.properties.pmoListTitle || 'PMO decisions',
            pmoIntegrationLookupName: this.properties.pmoIntegrationLookupName || 'Integration',
            stepsConfig: stepsConfig
        };
        var element = React.createElement(FormApp, props);
        ReactDom.render(element, this.domElement);
    };
    SmartFormWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(SmartFormWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    SmartFormWebPart.prototype.getPropertyPaneConfiguration = function () {
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
    };
    return SmartFormWebPart;
}(BaseClientSideWebPart));
export default SmartFormWebPart;
//# sourceMappingURL=SmartFormWebPart.js.map