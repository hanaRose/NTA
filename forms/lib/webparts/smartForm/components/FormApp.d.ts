import * as React from 'react';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { SPFI } from '@pnp/sp';
import '@pnp/sp/site-users/web';
import '@pnp/sp/views';
import '@pnp/sp/lists';
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