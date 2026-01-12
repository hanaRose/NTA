// src/webparts/smartForm/shared/constants.ts
export var SYSTEM_FIELD_NAMES = new Set([
    'Id', 'ID', 'Title', 'Author', 'AuthorId', 'Editor', 'EditorId',
    'Attachments', 'AttachmentFiles', 'FileRef', 'FileDirRef',
    'Created', 'Modified', 'ContentType', 'ContentTypeId',
    'OData__UIVersionString', 'GUID', 'FileSystemObjectType',
    'odata.type', 'odata.id', 'odata.etag', 'odata.editLink', 'odata.metadata',
    'OData__ColorTag', 'ComplianceAssetId', 'IntegrationId', 'ServerRedirectedEmbedUrl', 'ServerRedirectedEmbedUri', 'DateForIntegrationTeamDecisionIm', 'sentProtocol'
]);
export function isSystemField(name) {
    if (!name)
        return true;
    return SYSTEM_FIELD_NAMES.has(String(name));
}
// Helper to infer phase number from a value (string/number)
export function parsePhase(v) {
    if (v === null || v === undefined)
        return null;
    var s = String(v).trim().toLowerCase();
    if (s === '1' || s.indexOf('phase 1') > -1 || s.indexOf('pase 1') > -1)
        return 1;
    if (s === '2' || s.indexOf('phase 2') > -1 || s.indexOf('pase 2') > -1)
        return 2;
    if (s === '3' || s.indexOf('phase 3') > -1 || s.indexOf('pase 3') > -1)
        return 3;
    return null;
}
/*

// שדות מערכת/עזר שלא מציגים בטופס ולא שולחים בשמירה /
export const SYSTEM_FIELDS = new Set<string>([
  // OData/system
  'odata.type',
  'odata.id',
  'odata.etag',
  'odata.editLink',
  'odata.metadata',

  'FileSystemObjectType',
  'ContentTypeId',
  'Created',
  'Modified',
  'OData__UIVersionString',
  'GUID',

  // תוספות שנדרשו להסתיר
  'OData__ColorTag',
  'ComplianceAssetId', // "מזהה נכס תאימות"
  'IntegrationId',

  // גם אלו לא מציגים
  'Id',          // מזהה
  'AuthorId',    // Created By (id)
  'EditorId',    // Modified By (id)
  'Attachments', // דגל
  'AttachmentFiles' // אוסף קבצים מצורפים
]);

// true אם זה שדה מערכת או מתחיל ב-odata. /
export function isSystemField(internal: string): boolean {
  if (!internal) return true;
  if (SYSTEM_FIELDS.has(internal)) return true;
  if ((internal || '').toLowerCase().indexOf('odata.') === 0) return true;
  return false;
}

*/
//# sourceMappingURL=constants.js.map