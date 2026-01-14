import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
export interface ISmartFormWebPartProps {
    /** שם רשימת ה-PMO לעריכה */
    pmoListTitle: string;
    /** שם עמודת ה-Lookup ב-PMO שמקשרת ל-Integration */
    pmoIntegrationLookupName?: string;
    /** JSON של שלבים -> שמות שדות (Display/Internal) */
    stepsConfigJson?: string;
}
export default class SmartFormWebPart extends BaseClientSideWebPart<ISmartFormWebPartProps> {
    render(): void;
    protected onDispose(): void;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): {
        pages: {
            header: {
                description: string;
            };
            groups: {
                groupName: string;
                groupFields: (import("@microsoft/sp-webpart-base").IPropertyPaneField<import("@microsoft/sp-webpart-base").IPropertyPaneTextFieldProps> | import("@microsoft/sp-webpart-base").IPropertyPaneField<import("@microsoft/sp-webpart-base").IPropertyPaneLabelProps>)[];
            }[];
        }[];
    };
}
//# sourceMappingURL=SmartFormWebPart.d.ts.map