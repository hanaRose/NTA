// src/webparts/smartForm/shared/internalNames.ts
// GUID של Integration שסיפקת:
var INTEGRATION_LIST_ID = "2c962132-409d-4bf2-9440-3b3b6c7975a0";
export default INTEGRATION_LIST_ID;
export { INTEGRATION_LIST_ID };
// השדה ממנו נבנה התקציר ב-DD (HTML → טקסט רגיל, 30 תווים)
export var INTEGRATION_PREVIEW_FIELD_INTERNAL = 'ParticularsAndDescriptionOfPropo';
// שדות של Integration שביקשת להסתיר לגמרי
export var EXCLUDED_INTEGRATION_FIELDS = new Set([
    'Integration', 'DecisionRegardingProposedChange',
    'DecisionRegardingProposedChange_', 'DecisionappliestootherWorksTende',
    'IntegrationTeamDecision', 'RFCresponseLetterno',
    'RFCresponseaspublished', 'TenderCommitteeapprovaldate',
    'StatusofRFCresponse_x002f_TCRFC', 'Duedate', 'Actualdate',
    'OData__ColorTag', 'ComplianceAssetId', 'IntegrationId', 'IntegrationTeamDecisionEditingLa'
    // גם כל שדות מערכת יוסרו בפונקציית isSystemField()
]);
/*

// GUID של רשימת Integration (כפי שסיפקת) /
export const INTEGRATION_LIST_ID = '2c962132-409d-4bf2-9440-3b3b6c7975a0';
export const INTEGRATION_PREVIEW_FIELD_INTERNAL = 'ParticularsAndDescriptionOfPropo';


// שדות מתוך רשימת Integration שלא נציג בקריאה-בלבד (רועשים/לא נדרשים להצגה)
// (Internal Names שסיפקת; אפשר לעדכן/להרחיב לפי הצורך)

export const EXCLUDED_INTEGRATION_FIELDS = new Set<string>([
  // דוגמאות לשדות שנרצה להסתיר מקריאה-בלבד של Integration:
  'DecisionRegardingProposedChange',
  'DecisionRegardingProposedChange_',
  'DecisionappliestootherWorksTende',
  'RevisedWording_x002d_finalforpub',
  'RFC_x002f_TCRFCaspublishedbyNTA_',
  'IntegrationTeamDecision',
  'RelatedBidderRFCsexist_x003f_',
  'insertRelatedBidderRFCstexthere',
  'weretheRelatedBidderRFCsresponde',
  'insertRelatedBidderRFCresponse',
  'LMsProposedresponse',
  'RevisionincludeschangeinTenderDo',
  'RFCresponseLetterno',
  'RFCresponseaspublished',
  'Biddersnumber',
  'Assignedto',
  'Sub_x002d_Category',
  'TenderCommitteeapprovaldate',
  'StatusofRFCresponse_x002f_TCRFC',
  'Duedate',
  'Actualdate',
  // הוסף/הסירו לפי מה שתרצו שלא יוצג בחלק ה-Readonly
]);
*/ 
//# sourceMappingURL=internalNames.js.map