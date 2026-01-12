import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import type { IFieldInfo } from '@pnp/sp/fields/types';
export declare const FIELD_PERMISSION_LIST_TITLE = "fieldPermission";
export declare const FIELD_PERMISSION_LIST_ID = "05b813e8-e560-476d-89a7-5eac957bdc38";
export declare const FP_COL_INTERNAL = "internalFieldName";
export declare const FP_COL_WHO_CAN_EDIT = "WhoCanEdite";
export declare const GENERAL_ROLE_DEF_LIST_TITLE = "GeneralRoleDefinition";
export declare const COL_FINANCIAL_ADVISOR = "FinancialAdvisor";
export declare const COL_LAWYER = "Lawyer";
export declare const COL_PMO_INTEGRATION_TEAM = "PMOIntegrationTeam";
export declare const COL_PMO_TENDER_TEAM_FROM_CREATER = "formCreator";
export declare function getFieldMapsByTitle(sp: SPFI, listTitle: string): Promise<{
    internalToTitle: Record<string, string>;
    titleToInternal: Record<string, string>;
}>;
export declare function getFieldInfoMapById(sp: SPFI, listId: string): Promise<Record<string, IFieldInfo>>;
export declare function getFieldMapsById(sp: SPFI, listId: string): Promise<{
    internalToTitle: Record<string, string>;
    titleToInternal: Record<string, string>;
}>;
export declare function fetchIntegrationItemByGuid(sp: SPFI, integrationItemId: number | null): Promise<any>;
export declare function resolveInternalName(sp: SPFI, listTitle: string, displayOrInternal: string): Promise<string>;
export declare function getFieldInfoMap(sp: SPFI, listTitle: string): Promise<Record<string, IFieldInfo>>;
export declare function savePmoItem(sp: SPFI, pmoListId: string, id: number, draft: any): Promise<any>;
export declare const FP_COL_M1 = "OData__x004d_1";
export declare const FP_COL_M2 = "OData__x004d_2";
export declare const FP_COL_M3 = "OData__x004d_3";
export type FieldPermissionEntry = {
    default: string[];
    M1?: string[];
    M2?: string[];
    M3?: string[];
};
export type FieldPermissionMap = Record<string, FieldPermissionEntry>;
export declare function loadFieldPermissionMap(sp: SPFI, fieldPermissionListTitle?: string): Promise<FieldPermissionMap>;
export declare function fetchOrCreatePmoByIntegration(sp: SPFI, pmoListTitle: string, integrationItemId: number, integrationLookupDisplayOrInternal: string): Promise<{
    item: any;
    isNew: boolean;
}>;
export declare function loadGeneralRoleUsers(sp: SPFI): Promise<{
    FinancialAdvisor: string[];
    Lawyer: string[];
    PMOIntegrationTeam: string[];
    IntegrationTeam: string[];
    IntegrationTeamLawyer: string[];
    FinancialAdvisorIntegrationTeam: string[];
    M1TenderTeam: string[];
    LawyerM1TenderTeam: string[];
    M2TenderTeam: string[];
    LawyerM2TenderTeam: string[];
    M3TenderTeam: string[];
    LawyerM3TenderTeam: string[];
}>;
export declare function loadTenderTeamUsersFromIntegration(sp: SPFI, integrationItemId: number): Promise<string[]>;
export declare function canUserEditField(userEmailLc: string, internalName: string, fieldPermMap: FieldPermissionMap, roleUsers: {
    IntegrationTeam: string[];
    IntegrationTeamLawyer: string[];
    FinancialAdvisorIntegrationTeam: string[];
    M1TenderTeam: string[];
    LawyerM1TenderTeam: string[];
    M2TenderTeam: string[];
    LawyerM2TenderTeam: string[];
    M3TenderTeam: string[];
    LawyerM3TenderTeam: string[];
    FinancialAdvisor: string[];
    Lawyer: string[];
    PMOIntegrationTeam: string[];
    PMOTenderTeam: string[];
}, originatingLineManager?: string): boolean;
//# sourceMappingURL=data.d.ts.map