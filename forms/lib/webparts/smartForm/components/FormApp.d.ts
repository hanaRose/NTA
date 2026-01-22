import * as React from 'react';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { SPFI } from '@pnp/sp';
import '@pnp/sp/site-users/web';
import '@pnp/sp/views';
import '@pnp/sp/lists';
/**
 * 1) קוראת את כל המכרזים (Title) מרשימת SharePoint ומחזירה Set של שמות (trim).
 *    - עובדת גם אם יש יותר מ-5000 פריטים (ע"י paging).
 */
export declare function fetchAllTenderTitles(params: {
    sp: any;
    workTendersListId: string;
    titleFieldInternalName?: string;
}): Promise<Set<string>>;
/**
 * 2) בודקת שהמחרוזת היא:
 *    - בדיוק "All Infra 1 tenders"
 *    - או בדיוק "Not relevant to additional tenders"
 *    - או רשימת ערכים מופרדת בפסיקים, כשכל ערך קיים ברשימת המכרזים (Title)
 *    - בלי "עוד מילים" ובלי ערכים לא מוכרים.
 */
export declare function isValidTenderSelection(params: {
    input: string;
    validTitles: Set<string>;
    allowAllInfra?: boolean;
    allowNotRelevant?: boolean;
}): boolean;
export declare function buildPmoUpdatePayloadFromItem(pmoItem: any): any;
export declare function updateAutoCreatedPmoDecisionItem(params: {
    sp: any;
    integrationId: number;
    pmoItem: any;
    pmoLinkFieldInternalName: string;
    pmoListID?: string;
}): Promise<{
    ok: boolean;
    pmoDecisionId: any;
}>;
export declare function buildIntegrationPayloadFromPmo(integrationItem: any, pmoItem: any, pmoToIntegrationMap: Record<string, string>, extra?: Record<string, any>): any;
export declare function splitTenderAndCreateIntegrationItems(params: {
    sp: any;
    integrationListId: string;
    workTendersListId: string;
    pmoItem: any;
    itegrationItem: any;
    tenderSourceInternalName: string;
    integrationTenderInternalName?: string;
    pmoToIntegrationMap: Record<string, string>;
    linkFieldInternalName?: string;
    linkValue?: number;
    pmoDecisionsListId?: string;
    pmoIntegrationLookupIdField?: string;
    pmoDecisionAppliesFieldInternalName?: string;
    pmoSentProtocolFieldInternalName?: string;
    workTenderTitleField?: string;
    workTenderOlmField?: string;
    integrationOlmField?: string;
    decisionAppliesFieldInternalName?: string;
}): Promise<void>;
export type FieldInfoLike = {
    InternalName: string;
    TypeAsString?: string;
    Hidden?: boolean;
    ReadOnlyField?: boolean;
    Sealed?: boolean;
};
export declare function splitTenderAndCreateItemsFromTwoSources(params: {
    sp: any;
    listId: string;
    primaryItem: any;
    secondaryItem: any;
    fieldInfoMap: Record<string, FieldInfoLike>;
    tenderFieldInternalName?: string;
}): Promise<void>;
export declare function myOLM(sp: SPFI, generalRoleDefinitionListId: string, email: string): Promise<Array<"M1" | "M2" | "M3" | "ALL">>;
export interface FormAppProps {
    context: WebPartContext;
    pmoListTitle: string;
    pmoIntegrationLookupName?: string;
    stepsConfig?: Record<string, string[]>;
    isEditMode?: boolean;
}
declare const FormApp: React.FC<FormAppProps>;
export default FormApp;
//# sourceMappingURL=FormApp.d.ts.map