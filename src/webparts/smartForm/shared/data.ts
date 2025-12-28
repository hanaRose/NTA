

// src/webparts/smartForm/shared/data.ts
import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import type { IFieldInfo } from '@pnp/sp/fields/types';
//import { SPHttpClient } from '@microsoft/sp-http';
import { INTEGRATION_LIST_ID, EXCLUDED_INTEGRATION_FIELDS } from './internalNames';
import { isSystemField } from './constants';
import { spPost } from "@pnp/sp";

const lc = (s?: string) => String(s || '').toLowerCase();
//const strEq = (a?: string, b?: string) => lc(a) === lc(b);
//const strHas = (a?: string, frag?: string) => lc(a).indexOf(lc(frag)) > -1;

export const FIELD_PERMISSION_LIST_TITLE = 'fieldPermission';
export const FP_COL_INTERNAL = 'internalFieldName';   // Text
export const FP_COL_WHO_CAN_EDIT = 'WhoCanEdite';     // Choice/MultiChoice (טקסטים בעברית לפי התיאור)

export const GENERAL_ROLE_DEF_LIST_TITLE = 'GeneralRoleDefinition';
export const COL_FINANCIAL_ADVISOR = 'FinancialAdvisor';         // People (multi)
export const COL_LAWYER = 'Lawyer';                               // People (multi)
export const COL_PMO_INTEGRATION_TEAM = 'PMOIntegrationTeam';     // People (multi)

export const COL_PMO_TENDER_TEAM_FROM_CREATER = 'formCreator';    // People (multi) ברשימת Integration

function integrationList(sp: SPFI) {
  return sp.web.lists.getById(INTEGRATION_LIST_ID);
}

export async function getFieldMapsByTitle(
  sp: SPFI,
  listTitle: string
): Promise<{ internalToTitle: Record<string, string>; titleToInternal: Record<string, string> }> {
  const fields = await sp.web.lists.getByTitle(listTitle).fields
    .select('Title','InternalName')() as IFieldInfo[];

  const internalToTitle: Record<string, string> = {};
  const titleToInternal: Record<string, string> = {};

  for (let i = 0; i < fields.length; i++) {
    const f = fields[i];
    if (!f) continue;
    const disp = f.Title || f.InternalName;
    internalToTitle[f.InternalName] = disp;
    titleToInternal[disp] = f.InternalName;
  }
  return { internalToTitle, titleToInternal };
}

export async function getFieldInfoMapById(
  sp: SPFI,
  listId: string
): Promise<Record<string, IFieldInfo>> {
  const fields = await sp.web.lists
    .getById(listId)
    .fields
    .select(
      'InternalName',
      'Title',
      'TypeAsString',
      'ReadOnlyField',
      'Hidden',
      'RichText',
      'Choices',
      'AllowMultipleValues'
    )() as IFieldInfo[];

  const map: Record<string, IFieldInfo> = {};

  for (const f of fields) {
    if (f?.InternalName) {
      map[f.InternalName] = f;
    }
  }

  return map;
}




export async function getFieldMapsById(
  sp: SPFI,
  listId: string
): Promise<{ internalToTitle: Record<string, string>; titleToInternal: Record<string, string> }> {
  const fields = await sp.web.lists.getById(listId).fields
    .select('Title','InternalName')() as IFieldInfo[];

  const internalToTitle: Record<string, string> = {};
  const titleToInternal: Record<string, string> = {};

  for (let i = 0; i < fields.length; i++) {
    const f = fields[i];
    if (!f) continue;
    const disp = f.Title || f.InternalName;
    internalToTitle[f.InternalName] = disp;
    titleToInternal[disp] = f.InternalName;
  }
  return { internalToTitle, titleToInternal };
}

export async function fetchIntegrationItemByGuid(
  sp: SPFI,
  integrationItemId: number | null
): Promise<any> {
  console.log("🔮a");
  const list = integrationList(sp);
  console.log("🔮b");
  let item: any;
  if (integrationItemId) {
    console.log("🔮c");
    item = await list.items.getById(integrationItemId).select('*')();
  } else {
    console.log("🔮d");
    const arr = await list.items.orderBy('Id', false).top(1)();
    console.log("🔮e");
    item = arr && arr.length ? arr[0] : null;
    console.log("🔮f");
  }
  if (!item) return null;
  console.log("🔮g");
  const cleaned: any = {};
  console.log("🔮h");
  const keys = Object.keys(item);
  console.log("🔮i");
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    if (isSystemField(k)) continue;
    if (EXCLUDED_INTEGRATION_FIELDS.has(k)) continue;
    cleaned[k] = item[k];
  }
  console.log("🔮k");
  return cleaned;
}

export async function resolveInternalName(
  sp: SPFI,
  listTitle: string,
  displayOrInternal: string
): Promise<string> {
  const fields = await sp.web.lists.getByTitle(listTitle).fields
    .select('Title', 'InternalName')() as IFieldInfo[];

  let internal = displayOrInternal;
  const needleLc = lc(displayOrInternal);

  for (let i = 0; i < fields.length; i++) {
    const f = fields[i];
    if (!f) continue;
    const t = lc(f.Title || '');
    const n = lc(f.InternalName || '');
    if (n === needleLc || t === needleLc) { internal = f.InternalName; break; }
  }
  return internal;
}

export async function getFieldInfoMap(sp: SPFI, listTitle: string): Promise<Record<string, IFieldInfo>> {
  const fields = await sp.web.lists.getByTitle(listTitle).fields
    .select('InternalName','Title','TypeAsString','ReadOnlyField','Hidden','RichText','Choices','AllowMultipleValues')() as IFieldInfo[];
  const map: Record<string, IFieldInfo> = {};
  for (let i = 0; i < fields.length; i++) {
    const f = fields[i];
    if (f && f.InternalName) map[f.InternalName] = f;
  }
  return map;
}
/*
function isMultiField(info?: IFieldInfo): boolean {
  if (!info) return false;
  if ((info as any).AllowMultipleValues === true) return true;
  return strHas(info?.TypeAsString, 'multi');
}
function isLookupOrUser(info?: IFieldInfo): boolean {
  if (!info || !info.TypeAsString) return false;
  const t = lc(info.TypeAsString);
  return t.indexOf('lookup') > -1 || t.indexOf('user') > -1;
}
function isChoice(info?: IFieldInfo): boolean {
  if (!info || !info.TypeAsString) return false;
  const t = lc(info.TypeAsString);
  return t === 'choice' || t === 'multichoice';
}
function isNumberField(info?: IFieldInfo): boolean {
  if (!info || !info.TypeAsString) return false;
  const t = lc(info.TypeAsString);
  return t === 'number' || t === 'currency';
}
function isBooleanField(info?: IFieldInfo): boolean {
  if (!info || !info.TypeAsString) return false;
  return lc(info.TypeAsString) === 'boolean';
}

function normalizeNumber(v: any): number | null {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number' && isFinite(v)) return v;
  if (typeof v === 'string') {
    const s = v.trim();
    if (!s) return null;
    const normalized = s.replace(',', '.');
    const num = parseFloat(normalized);
    return isFinite(num) ? num : null;
  }
  return null;
}
function normalizeBoolean(v: any): boolean | null {
  if (v === null || v === undefined) return null;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (typeof v === 'string') {
    const s = v.trim().toLowerCase();
    if (!s) return null;
    if (s === 'true' || s === 'yes' || s === 'y' || s === '1') return true;
    if (s === 'false' || s === 'no' || s === 'n' || s === '0') return false;
  }
  return null;
}*/
/*
function buildCleanUpdatePayload(
  draft: any,
  fieldMap: Record<string, IFieldInfo>
): Record<string, any> {
  const out: Record<string, any> = {};

  const keys = Object.keys(draft || {});
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    if (!k || k === 'undefined') continue;

    const info = fieldMap[k];
    if (!info) continue;

    if ((info as any).Hidden || (info as any).ReadOnlyField) continue;
    if (strEq(k, 'Integration') || strEq(k, 'IntegrationId')) continue;

    const v = draft[k];

    // Lookup/User
    if (isLookupOrUser(info)) {
      const target = `${(info as any).InternalName}Id`;
      if (isMultiField(info)) {
        if (Array.isArray(v)) {
          const nums: number[] = [];
          for (let j = 0; j < v.length; j++) if (typeof v[j] === 'number') nums.push(v[j]);
          out[target] = { results: nums };
        } else if (typeof v === 'number') {
          out[target] = { results: [v] };
        }
      } else {
        if (typeof v === 'number') out[target] = v;
        else if (Array.isArray(v) && v.length && typeof v[0] === 'number') out[target] = v[0];
      }
      continue;
    }

    // Choice/MultiChoice
    if (isChoice(info)) {
      if (isMultiField(info)) {
        if (Array.isArray(v)) {
          const arr: string[] = [];
          for (let j = 0; j < v.length; j++) arr.push(String(v[j]));
          out[(info as any).InternalName] = arr;
        } else if (typeof v === 'string' && v) {
          out[(info as any).InternalName] = [v];
        }
      } else {
        if (typeof v === 'string' || v === null) out[(info as any).InternalName] = v;
      }
      continue;
    }

    // Date/DateTime
    if (strHas((info as any).TypeAsString, 'datetime') || strHas((info as any).TypeAsString, 'date')) {
      if (v instanceof Date) out[(info as any).InternalName] = v.toISOString();
      else if (typeof v === 'string' || v === null) out[(info as any).InternalName] = v;
      continue;
    }

    // Number/Currency
    if (isNumberField(info)) {
      const num = normalizeNumber(v);
      out[(info as any).InternalName] = num;
      continue;
    }

    // Boolean
    if (isBooleanField(info)) {
      const b = normalizeBoolean(v);
      out[(info as any).InternalName] = b;
      continue;
    }

    // Default primitives
    if (v === null || typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') {
      out[(info as any).InternalName] = v;
      continue;
    }
  }

  return out;
}*/
function cleanRichText(html: string | null | undefined): string | null {
  if (!html) return null;  // במקום להחזיר undefined — מחזירים null תמיד

  let cleaned: string = html;

  cleaned = cleaned.replace(/class="ExternalClass[^"]*"/g, "");
  cleaned = cleaned.replace(/class="editor-paragraph"/g, "");
  cleaned = cleaned.replace(/<div>\s*<\/div>/g, "");
  cleaned = cleaned.replace(/<p>\s*<\/p>/g, "");
  cleaned = cleaned.replace(/<div>\s*<div>/g, "<div>");
  cleaned = cleaned.replace(/<\/div>\s*<\/div>/g, "</div>");

  return cleaned.trim();
}

function isValidHttpUrl(raw: string): boolean {
  if (!raw) return false;
  try {
    const u = new URL(raw.trim());
    return u.protocol === "http:" || u.protocol === "https:";
  } catch {
    return false;
  }
}


export async function savePmoItem(
  sp: SPFI,
  pmoListId: string,
  id: number,
  draft: any
): Promise<any> {
  console.log("savePmoItem 🌭");
  const list = sp.web.lists.getById(pmoListId);

  // 1. מעתיקים את האובייקט כדי לא לשנות את המקור
  const payload: any = { ...(draft || {}) };

  // 2. זורקים שדות מערכת שלא צריך לשלוח ל־SharePoint
  delete payload.ServerRedirectedEmbedUri;
  delete payload.ServerRedirectedEmbedUrl;
  delete payload.Id;
  delete payload.ID;
  delete payload.odata;
  delete payload["OData__ColorTag"];
  delete payload['odata.type'];
  delete payload['odata.id'];
  delete payload['odata.etag'];
  delete payload['odata.editLink'];
  delete payload.ComplianceAssetId;
  delete payload.FileSystemObjectType;
  delete payload.OData__UIVersionString;
  delete payload.AuthorId;
  delete payload.EditorId;
  delete payload.Created;
  delete payload.Modified;
  delete payload.RFCorTcRFCasPublishedByNTaToBeFi;
  delete payload.RFCorTcRFCasPublishedByNTaNew;

  payload.RevisedWordingFinalForPublicatio =
    cleanRichText(payload.RevisedWordingFinalForPublicatio);


  console.log('🧾 payload sent to SharePoint:', payload);
  console.log('🐶 dog in payload:', payload.dog);

  // GET ה־etag כדי SharePoint יאפשר עדכון
for (const key in payload) {
  const val = payload[key];
  if (!payload.RFCResponseLetterNo && payload.RFCResponseLetterNo !== 0) {
    delete payload.RFCResponseLetterNo;
  }

  if (payload.DecisionDate) {
    const base = new Date(payload.DecisionDate);
    const plus14 = new Date(base.getTime() + 14 * 24 * 60 * 60 * 1000);
    payload.DueDateCalculated = plus14.toISOString();
  }

  /*
  // Hyperlink field?
  if (val && typeof val === "object" && ("Url" in val || "url" in val)) {
    const url = (val.Url || val.url || "").trim();
    const desc = (val.Description || val.description || url).trim();

    if (url) {
      payload[key] = {
        "__metadata": { "type": "SP.FieldUrlValue" },
        "Url": url,
        "Description": desc
      };
    } else {
      // אין URL → מאפסים
      payload[key] = null;
    }
  }*/
 
   // Hyperlink field?
  if (val && typeof val === "object" && ("Url" in val || "url" in val)) {
    const url = (val.Url || val.url || "").trim();
    const desc = (val.Description || val.description || url).trim();

    if (!url) {
      // אין URL → מאפסים
      payload[key] = null;
    } else {
      // ✅ בדיקת תקינות ה-URL
      if (!isValidHttpUrl(url)) {
        // כאן זו הטלת שגיאה שתיתפס ב-try/catch שקורא ל-savePmoItem
        throw new Error("the url entered is not valid");
      }

      payload[key] = {
        "__metadata": { "type": "SP.FieldUrlValue" },
        "Url": url,
        "Description": desc
      };
    }
  }


  // MultiChoice field fix
  if (Array.isArray(val)) {
    payload[key] = {
      "__metadata": { "type": "Collection(Edm.String)" },
      "results": val
    };
  }


}
//const webUrl = this.context.pageContext.web.absoluteUrl;
console.log("savePmoItem 🌭🌭");
console.log("SAVE PAYLOAD", JSON.stringify(payload, null, 2));
payload.__metadata = {
  type: "SP.Data.PMO_x0020_decisionsListItem"
};
// ⭐ UPDATE ללא metadata — רק REST
  await spPost(
    list.items.getById(id),
    {
      headers: {
        "Accept": "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=verbose",
        "IF-MATCH": "*",
        "X-HTTP-Method": "MERGE"
      },
      body: JSON.stringify(payload)
    }
  );


console.log("savePmoItem 🌭🌭🌭");
  // 3. עדכון בפועל
  //await list.items.getById(id).update(payload);

  // 4. מחזירים את הפריט המעודכן
  return await list.items.getById(id).select('*')();
}

/*

export async function savePmoItem(
  sp: SPFI,
  pmoListId: string,   // <-- עכשיו זה ID ולא Title
  id: number,
  draft: any
): Promise<any> {
  const list = sp.web.lists.getById(pmoListId);
  const fieldMap = await getFieldInfoMapById(sp, pmoListId);
  const clean = buildCleanUpdatePayload(draft, fieldMap);

  await list.items.getById(id).update(clean);
  return await list.items.getById(id).select('*')();
}*/

/*
export async function savePmoItem(
  sp: SPFI,
  pmoListTitle: string,
  id: number,
  draft: any
): Promise<any> {
  const list = sp.web.lists.getByTitle(pmoListTitle);
  const fieldMap = await getFieldInfoMap(sp, pmoListTitle);
  const clean = buildCleanUpdatePayload(draft, fieldMap);

  await list.items.getById(id).update(clean);
  const saved = await list.items.getById(id).select('*')();
  return saved;
}*/

// =========================================================================
//   ==  הרשאות שדה ברמת PMO decisions לפי fieldPermission + קבוצות ==
//   ========================================================================= 


// 🔹 עמודות חדשות לשדות לפי M1/M2/M3 ברשימת fieldPermission
export const FP_COL_M1 = 'OData__x004d_1';
export const FP_COL_M2 = 'OData__x004d_2';
export const FP_COL_M3 = 'OData__x004d_3';

// 🔹 טיפוס מורכב לכל שדה: ברירת מחדל + ספציפי ל-M1 / M2 / M3
export type FieldPermissionEntry = {
  default: string[];   // מהעמודה WhoCanEdite
  M1?: string[];
  M2?: string[];
  M3?: string[];
};

// 🔹 המפה הכוללת: internalName -> Entry
export type FieldPermissionMap = Record<string, FieldPermissionEntry>;

// טוען מפה: internalFieldName -> מערך תפקידי הרשאה (בחירות מ- WhoCanEdite) 
// טוען מפה: internalFieldName -> אובייקט הרשאות (ברירת מחדל + M1/M2/M3)
export async function loadFieldPermissionMap(
  sp: SPFI,
  fieldPermissionListTitle: string = FIELD_PERMISSION_LIST_TITLE
): Promise<FieldPermissionMap> {
  console.log("😋1_ fieldPermissionListTitle ", fieldPermissionListTitle);
  const list = sp.web.lists.getByTitle(fieldPermissionListTitle);
  console.log("😋2_ FP_COL_INTERNAL ", FP_COL_INTERNAL, "FP_COL_WHO_CAN_EDIT ", FP_COL_WHO_CAN_EDIT);

  // שימי לב: מוסיפים גם את העמודות M1 / M2 / M3 ל-select
  const items = await list.items.select('*')();//('Id', FP_COL_INTERNAL, FP_COL_WHO_CAN_EDIT, FP_COL_M1, FP_COL_M2, FP_COL_M3)();

  console.log("😋3_ ✏️ fieldPermission items: ", items);

  const map: FieldPermissionMap = {};

  // פונקציית עזר – ממירה ערך SharePoint (MultiChoice/Single/Text) למערך מחרוזות
  const toChoices = (raw: any): string[] => {
    console.log("😋 4_");
    if (!raw) return [];
    if (Array.isArray(raw)) return raw.map(x => String(x));
    if (raw && Array.isArray(raw.results)) return (raw.results as any[]).map(x => String(x));
    if (typeof raw === 'string' && raw.trim()) return [raw.trim()];
    return [];
  };

  for (let i = 0; i < items.length; i++) {
   
    const it = items[i] || {};
    const internal = it[FP_COL_INTERNAL];
    if (!internal) continue;

    const base = toChoices(it[FP_COL_WHO_CAN_EDIT]);
    const m1 = toChoices(it[FP_COL_M1]);
    const m2 = toChoices(it[FP_COL_M2]);
    const m3 = toChoices(it[FP_COL_M3]);

    // אם כבר יש רשומה לשדה הזה – נאחד
    let entry: FieldPermissionEntry = map[internal] || { default: [] };

    const pushDistinctAll = (target: string[], src: string[]) => {
      for (let j = 0; j < src.length; j++) {
        const v = String(src[j]);
        if (target.indexOf(v) === -1) target.push(v);
      }
    };

    pushDistinctAll(entry.default, base);
    if (m1.length) {
      entry.M1 = entry.M1 || [];
      pushDistinctAll(entry.M1, m1);
    }
    if (m2.length) {
      entry.M2 = entry.M2 || [];
      pushDistinctAll(entry.M2, m2);
    }
    if (m3.length) {
      entry.M3 = entry.M3 || [];
      pushDistinctAll(entry.M3, m3);
    }

    map[internal] = entry;
  }
   console.log("😋 5_");
  return map;
}


// ✅ ייצוא בשם מהמקור
export async function fetchOrCreatePmoByIntegration(
  sp: SPFI,
  pmoListTitle: string,
  integrationItemId: number,
  integrationLookupDisplayOrInternal: string
): Promise<{ item: any; isNew: boolean }> {
  console.log("🔮1a");
  //const list = sp.web.lists.getByTitle(pmoListTitle);
  const list = sp.web.lists.getById('e5e8eaea-16db-49d3-ad7c-62f5a2bdd97a');
  console.log("🔮2b");
  // resolve lookup internal name
  let fields;
  try{
  fields = await list.fields.select('Title','InternalName')();///////////////////////////////////////////////
  } catch (e:any) {
  console.warn("🔮", e);
  // fallback: leave lookupInternal as-is
}
  console.log("🔮3d");
  let lookupInternal = integrationLookupDisplayOrInternal;
  console.log("🔮4r")
  for (const f of fields as any[]) {
    const wantLc  = integrationLookupDisplayOrInternal.toLowerCase();
    if (String(f.Title || '').toLowerCase() === wantLc ||
        String(f.InternalName || '').toLowerCase() === wantLc) {
      lookupInternal = f.InternalName;
      break;
    }
  }
  console.log("🔮5e");
  const found = await list.items.filter(`${lookupInternal}Id eq ${integrationItemId}`).top(1)();
  console.log("🔮6r");
  if (found && found.length) return { item: found[0], isNew: false };
  console.log("🔮7g");
  const created = await list.items.add({ [`${lookupInternal}Id`]: integrationItemId });
  console.log("🔮8g");
 const item = await list.items.getById(created.data.Id).select('*')();
 console.log("🔮9j");
  return { item, isNew: true };
}












const COL_INTEGRATION_TEAM = 'IntegrationTeam';
const COL_INTEGRATION_TEAM_LAWYER = 'IntegrationTeamLawyer';
const COL_FINANCIAL_ADVISOR_INTEGRATION_TEAM = 'FinancialAdvisorIntegrationTeam';
const COL_M1_TENDER_TEAM = 'M1TenderTeam';
const COL_LAWYER_M1_TENDER_TEAM = 'LawyerM1TenderTeam';
const COL_M2_TENDER_TEAM = 'M2TenderTeam';
const COL_LAWYER_M2_TENDER_TEAM = 'LawyerM2TenderTeam';
const COL_M3_TENDER_TEAM = 'M3TenderTeam';
const COL_LAWYER_M3_TENDER_TEAM = 'LawyerM3TenderTeam';


// שליפת כל המשתמשים ברשימת GeneralRoleDefinition לשלוש קבוצות + 9 עמודות טקסט
export async function loadGeneralRoleUsers(
  sp: SPFI
): Promise<{
  FinancialAdvisor: string[]; // emails lower-case
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
}> {
  const list = sp.web.lists.getByTitle(GENERAL_ROLE_DEF_LIST_TITLE);

  const rows = await list.items
    .select(
      'Id',
      `${COL_FINANCIAL_ADVISOR}/EMail`,
      `${COL_LAWYER}/EMail`,
      `${COL_PMO_INTEGRATION_TEAM}/EMail`,
      COL_INTEGRATION_TEAM,
      COL_INTEGRATION_TEAM_LAWYER,
      COL_FINANCIAL_ADVISOR_INTEGRATION_TEAM,
      COL_M1_TENDER_TEAM,
      COL_LAWYER_M1_TENDER_TEAM,
      COL_M2_TENDER_TEAM,
      COL_LAWYER_M2_TENDER_TEAM,
      COL_M3_TENDER_TEAM,
      COL_LAWYER_M3_TENDER_TEAM
    )
    .expand(COL_FINANCIAL_ADVISOR, COL_LAWYER, COL_PMO_INTEGRATION_TEAM)();

  const fin: string[] = [];
  const law: string[] = [];
  const pmo: string[] = [];

  const integrationTeam: string[] = [];
  const integrationTeamLawyer: string[] = [];
  const financialAdvisorIntegrationTeam: string[] = [];
  const m1TenderTeam: string[] = [];
  const lawyerM1TenderTeam: string[] = [];
  const m2TenderTeam: string[] = [];
  const lawyerM2TenderTeam: string[] = [];
  const m3TenderTeam: string[] = [];
  const lawyerM3TenderTeam: string[] = [];

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || {};

    // PEOPLE columns → מערכים של אובייקטי משתמש
    const addPeople = (arr: any, target: string[]) => {
      if (!arr) return;
      const src = Array.isArray(arr) ? arr : (arr.results || []);
      for (let j = 0; j < src.length; j++) {
        const p = src[j];
        const email = lc(p && (p.EMail || p.Email || p.UserPrincipalName));
        if (email && target.indexOf(email) === -1) target.push(email);
      }
    };

    // TEXT columns → מפצלים למיילים (אם כתבְת שם כמה מיילים מופרדים בפסיק / נקודה־פסיק / רווחים)
    const addTextEmails = (val: any, target: string[]) => {
      if (!val) return;
      const parts = String(val).split(/[;,\s]+/);
      for (let k = 0; k < parts.length; k++) {
        const email = lc(parts[k]);
        if (email && target.indexOf(email) === -1) {
          target.push(email);
        }
      }
    };

    // PEOPLE
    addPeople(r[COL_FINANCIAL_ADVISOR], fin);
    addPeople(r[COL_LAWYER], law);
    addPeople(r[COL_PMO_INTEGRATION_TEAM], pmo);

    // TEXT (9 העמודות החדשות)
    addTextEmails(r[COL_INTEGRATION_TEAM], integrationTeam);
    addTextEmails(r[COL_INTEGRATION_TEAM_LAWYER], integrationTeamLawyer);
    addTextEmails(r[COL_FINANCIAL_ADVISOR_INTEGRATION_TEAM], financialAdvisorIntegrationTeam);
    addTextEmails(r[COL_M1_TENDER_TEAM], m1TenderTeam);
    addTextEmails(r[COL_LAWYER_M1_TENDER_TEAM], lawyerM1TenderTeam);
    addTextEmails(r[COL_M2_TENDER_TEAM], m2TenderTeam);
    addTextEmails(r[COL_LAWYER_M2_TENDER_TEAM], lawyerM2TenderTeam);
    addTextEmails(r[COL_M3_TENDER_TEAM], m3TenderTeam);
    addTextEmails(r[COL_LAWYER_M3_TENDER_TEAM], lawyerM3TenderTeam);
  }

  return {
    IntegrationTeam: integrationTeam,
    IntegrationTeamLawyer: integrationTeamLawyer,
    FinancialAdvisorIntegrationTeam: financialAdvisorIntegrationTeam,
    M1TenderTeam: m1TenderTeam,
    LawyerM1TenderTeam: lawyerM1TenderTeam,
    M2TenderTeam: m2TenderTeam,
    LawyerM2TenderTeam: lawyerM2TenderTeam,
    M3TenderTeam: m3TenderTeam,
    LawyerM3TenderTeam: lawyerM3TenderTeam,

    FinancialAdvisor: fin,
    Lawyer: law,
    PMOIntegrationTeam: pmo,
  };
}





export async function loadTenderTeamUsersFromIntegration(
  sp: SPFI,
  integrationItemId: number
): Promise<string[]> {
  // שליפה של השדה כטקסט – ללא expand
  console.log("🫙🔪🎡🎊🎀🎁🎉✨🧨🎆🎈 in loadTenderTeamUsersFromIntegration");
  const it = await integrationList(sp).items
    .getById(integrationItemId)
    .select('Id', COL_PMO_TENDER_TEAM_FROM_CREATER)();
  console.log("it ", it);
  // אם השדה ריק – נחזיר ריק
  const raw = String(it?.[COL_PMO_TENDER_TEAM_FROM_CREATER] || '');
  console.log("🪂 COL_PMO_TENDER_TEAM_FROM_CREATER ", COL_PMO_TENDER_TEAM_FROM_CREATER, "raw ", raw);
  // מאפשר פסיקים/נקודה-פסיק/שורות מרובות
  const parts = raw.split(/[,;]|\r?\n/);
  console.log("🎈parts ", parts);
  const emails: string[] = [];
  for (let i = 0; i < parts.length; i++) {
    const e = parts[i].trim().toLowerCase();
    if (!e) continue;
    // ולידציה בסיסית של אימייל + מניעת כפילויות
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e) && emails.indexOf(e) === -1) {
      emails.push(e);
    }
  }
  console.log("📧 emails ", emails); 
  return emails;
}


// החלטת עריכה לשדה ספציפי עבור משתמש נתון (email) /
// החלטת עריכה לשדה ספציפי עבור משתמש נתון (email) /
// כעת: לוקחת בחשבון גם OriginatingLineManager (M1/M2/M3)
export function canUserEditField(
  userEmailLc: string,
  internalName: string,
  fieldPermMap: FieldPermissionMap,
  roleUsers: {
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
    PMOTenderTeam: string[]; // מגיע מ-Integration.FromCreater
  },
  originatingLineManager?: string  // 🔹 חדש: 'M1' | 'M2' | 'M3'
): boolean {
  const entry = fieldPermMap[internalName];
  if (!entry) {
    console.log("🧨🛼internalName ", internalName, " entry ", entry);
    return false;} // אין הרשאות מוגדרות לשדה → קריאה בלבד

  // בחירת סט ה-"בחירות" לפי OriginatingLineManager:
  const olm = String(originatingLineManager || '').trim().toUpperCase();

  let choices: string[] = entry.default || [];

  if (olm === 'M1' && entry.M1 && entry.M1.length) {
    choices = entry.M1;
  } else if (olm === 'M2' && entry.M2 && entry.M2.length) {
    choices = entry.M2;
  } else if (olm === 'M3' && entry.M3 && entry.M3.length) {
    choices = entry.M3;
  }
  console.log("🧨 choices ", choices);
  if (!choices || !choices.length) {
    console.log("🧨🧨🧨🧨🧨🧨 !choices || !choices.length ")
    // אם גם ב-default וגם ב-M1/M2/M3 אין כלום → אין הרשאה
    return false;
  }

  // בניית איחוד משתמשים לפי הבחירות
  const allowEmails: string[] = [];
  const pushDistinct = (src: string[]) => {
    for (let i = 0; i < src.length; i++) {
      const e = src[i];
      if (allowEmails.indexOf(e) === -1) allowEmails.push(e);
    }
  };

  for (let i = 0; i < choices.length; i++) {
    const choice = choices[i];
    const cLc = lc(choice);

    if (cLc.indexOf(lc('צוות אינטגרציה')) > -1) {pushDistinct(roleUsers.IntegrationTeam);console.log("🧨🧨צוות אינטגרציה'");}
    else if (cLc.indexOf(lc('עו"ד צוות אינטגרציה')) > -1) {pushDistinct(roleUsers.IntegrationTeamLawyer);console.log("🧨🧨עוד צוות אינטגרציה");}
    else if (cLc.indexOf(lc('יועץ פיננסי צוות אינטגרציה')) > -1) {pushDistinct(roleUsers.FinancialAdvisorIntegrationTeam);console.log("🧨🧨יועץ פיננסי צוות אינטגרציה'")}
    else if (cLc.indexOf(lc('צוות מכרז M1')) > -1) {pushDistinct(roleUsers.M1TenderTeam);console.log("🧨🧨'צוות מכרז M1'");}
    else if (cLc.indexOf(lc('עו"ד צוות מכרז M1')) > -1){ pushDistinct(roleUsers.LawyerM1TenderTeam);console.log("🧨🧨עוד צוות מכרז M1");}
    else if (cLc.indexOf(lc('צוות מכרז M2')) > -1) {pushDistinct(roleUsers.M2TenderTeam);console.log("🧨🧨צוות מכרז M2 roleUsers.M2TenderTeam ", roleUsers.M2TenderTeam);}
    else if (cLc.indexOf(lc('עו"ד צוות מכרז M2')) > -1) {pushDistinct(roleUsers.LawyerM2TenderTeam);console.log("🧨🧨צוות מכרז עוד M2");}
    else if (cLc.indexOf(lc('צוות מכרז M3')) > -1) {pushDistinct(roleUsers.M3TenderTeam);console.log("🧨🧨צוות מכרז M3");}
    else if (cLc.indexOf(lc('עו"ד צוות מכרז M3')) > -1) {pushDistinct(roleUsers.LawyerM3TenderTeam);console.log("🧨🧨צוות מכרז עוד M3");}
    // כאן אפשר להוסיף עוד מיפויים אם יהיו בחירות נוספות בטבלה
  }
  console.log("🧨😶‍🌫️allowEmails  ", allowEmails);
  console.log(" return allowEmails.indexOf(userEmailLc) > -1 ", allowEmails.indexOf(userEmailLc) > -1);
  return allowEmails.indexOf(userEmailLc) > -1;
}


/*

// src/webparts/smartForm/shared/data.ts
import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import type { IFieldInfo } from '@pnp/sp/fields/types';

import { INTEGRATION_LIST_ID, EXCLUDED_INTEGRATION_FIELDS } from './internalNames';
import { isSystemField } from './constants';

const lc = (s?: string) => String(s || '').toLowerCase();
const strEq = (a?: string, b?: string) => lc(a) === lc(b);
const strHas = (a?: string, frag?: string) => lc(a).indexOf(lc(frag)) > -1;

export const FIELD_PERMISSION_LIST_TITLE = 'fieldPermission';
export const FP_COL_INTERNAL = 'internalFieldName';   // Text
export const FP_COL_WHO_CAN_EDIT = 'WhoCanEdite';     // Choice/MultiChoice (טקסטים בעברית לפי התיאור)

export const GENERAL_ROLE_DEF_LIST_TITLE = 'GeneralRoleDefinition';
export const COL_FINANCIAL_ADVISOR = 'FinancialAdvisor';         // People (multi)
export const COL_LAWYER = 'Lawyer';                               // People (multi)
export const COL_PMO_INTEGRATION_TEAM = 'PMOIntegrationTeam';     // People (multi)

export const COL_PMO_TENDER_TEAM_FROM_CREATER = 'formCreator';    // People (multi) ברשימת Integration

function integrationList(sp: SPFI) {
  return sp.web.lists.getById(INTEGRATION_LIST_ID);
}

export async function getFieldMapsByTitle(
  sp: SPFI,
  listTitle: string
): Promise<{ internalToTitle: Record<string, string>; titleToInternal: Record<string, string> }> {
  const fields = await sp.web.lists.getByTitle(listTitle).fields
    .select('Title','InternalName')() as IFieldInfo[];

  const internalToTitle: Record<string, string> = {};
  const titleToInternal: Record<string, string> = {};

  for (let i = 0; i < fields.length; i++) {
    const f = fields[i];
    if (!f) continue;
    const disp = f.Title || f.InternalName;
    internalToTitle[f.InternalName] = disp;
    titleToInternal[disp] = f.InternalName;
  }
  return { internalToTitle, titleToInternal };
}

export async function getFieldMapsById(
  sp: SPFI,
  listId: string
): Promise<{ internalToTitle: Record<string, string>; titleToInternal: Record<string, string> }> {
  const fields = await sp.web.lists.getById(listId).fields
    .select('Title','InternalName')() as IFieldInfo[];

  const internalToTitle: Record<string, string> = {};
  const titleToInternal: Record<string, string> = {};

  for (let i = 0; i < fields.length; i++) {
    const f = fields[i];
    if (!f) continue;
    const disp = f.Title || f.InternalName;
    internalToTitle[f.InternalName] = disp;
    titleToInternal[disp] = f.InternalName;
  }
  return { internalToTitle, titleToInternal };
}

export async function fetchIntegrationItemByGuid(
  sp: SPFI,
  integrationItemId: number | null
): Promise<any> {
  const list = integrationList(sp);

  let item: any;
  if (integrationItemId) {
    item = await list.items.getById(integrationItemId).select('*')();
  } else {
    const arr = await list.items.orderBy('Id', false).top(1)();
    item = arr && arr.length ? arr[0] : null;
  }
  if (!item) return null;

  const cleaned: any = {};
  const keys = Object.keys(item);
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    if (isSystemField(k)) continue;
    if (EXCLUDED_INTEGRATION_FIELDS.has(k)) continue;
    cleaned[k] = item[k];
  }
  return cleaned;
}

export async function resolveInternalName(
  sp: SPFI,
  listTitle: string,
  displayOrInternal: string
): Promise<string> {
  const fields = await sp.web.lists.getByTitle(listTitle).fields
    .select('Title', 'InternalName')() as IFieldInfo[];

  let internal = displayOrInternal;
  const needleLc = lc(displayOrInternal);

  for (let i = 0; i < fields.length; i++) {
    const f = fields[i];
    if (!f) continue;
    const t = lc(f.Title || '');
    const n = lc(f.InternalName || '');
    if (n === needleLc || t === needleLc) { internal = f.InternalName; break; }
  }
  return internal;
}

export async function getFieldInfoMap(sp: SPFI, listTitle: string): Promise<Record<string, IFieldInfo>> {
  const fields = await sp.web.lists.getByTitle(listTitle).fields
    .select('InternalName','Title','TypeAsString','ReadOnlyField','Hidden','RichText','Choices','AllowMultipleValues')() as IFieldInfo[];
  const map: Record<string, IFieldInfo> = {};
  for (let i = 0; i < fields.length; i++) {
    const f = fields[i];
    if (f && f.InternalName) map[f.InternalName] = f;
  }
  return map;
}

function isMultiField(info?: IFieldInfo): boolean {
  if (!info) return false;
  if ((info as any).AllowMultipleValues === true) return true;
  return strHas(info?.TypeAsString, 'multi');
}
function isLookupOrUser(info?: IFieldInfo): boolean {
  if (!info || !info.TypeAsString) return false;
  const t = lc(info.TypeAsString);
  return t.indexOf('lookup') > -1 || t.indexOf('user') > -1;
}
function isChoice(info?: IFieldInfo): boolean {
  if (!info || !info.TypeAsString) return false;
  const t = lc(info.TypeAsString);
  return t === 'choice' || t === 'multichoice';
}
function isNumberField(info?: IFieldInfo): boolean {
  if (!info || !info.TypeAsString) return false;
  const t = lc(info.TypeAsString);
  return t === 'number' || t === 'currency';
}
function isBooleanField(info?: IFieldInfo): boolean {
  if (!info || !info.TypeAsString) return false;
  return lc(info.TypeAsString) === 'boolean';
}

function normalizeNumber(v: any): number | null {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number' && isFinite(v)) return v;
  if (typeof v === 'string') {
    const s = v.trim();
    if (!s) return null;
    const normalized = s.replace(',', '.');
    const num = parseFloat(normalized);
    return isFinite(num) ? num : null;
  }
  return null;
}
function normalizeBoolean(v: any): boolean | null {
  if (v === null || v === undefined) return null;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (typeof v === 'string') {
    const s = v.trim().toLowerCase();
    if (!s) return null;
    if (s === 'true' || s === 'yes' || s === 'y' || s === '1') return true;
    if (s === 'false' || s === 'no' || s === 'n' || s === '0') return false;
  }
  return null;
}

function buildCleanUpdatePayload(
  draft: any,
  fieldMap: Record<string, IFieldInfo>
): Record<string, any> {
  const out: Record<string, any> = {};

  const keys = Object.keys(draft || {});
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    if (!k || k === 'undefined') continue;

    const info = fieldMap[k];
    if (!info) continue;

    if ((info as any).Hidden || (info as any).ReadOnlyField) continue;
    if (strEq(k, 'Integration') || strEq(k, 'IntegrationId')) continue;

    const v = draft[k];

    // Lookup/User
    if (isLookupOrUser(info)) {
      const target = `${(info as any).InternalName}Id`;
      if (isMultiField(info)) {
        if (Array.isArray(v)) {
          const nums: number[] = [];
          for (let j = 0; j < v.length; j++) if (typeof v[j] === 'number') nums.push(v[j]);
          out[target] = { results: nums };
        } else if (typeof v === 'number') {
          out[target] = { results: [v] };
        }
      } else {
        if (typeof v === 'number') out[target] = v;
        else if (Array.isArray(v) && v.length && typeof v[0] === 'number') out[target] = v[0];
      }
      continue;
    }

    // Choice/MultiChoice
    if (isChoice(info)) {
      if (isMultiField(info)) {
        if (Array.isArray(v)) {
          const arr: string[] = [];
          for (let j = 0; j < v.length; j++) arr.push(String(v[j]));
          out[(info as any).InternalName] = arr;
        } else if (typeof v === 'string' && v) {
          out[(info as any).InternalName] = [v];
        }
      } else {
        if (typeof v === 'string' || v === null) out[(info as any).InternalName] = v;
      }
      continue;
    }

    // Date/DateTime
    if (strHas((info as any).TypeAsString, 'datetime') || strHas((info as any).TypeAsString, 'date')) {
      if (v instanceof Date) out[(info as any).InternalName] = v.toISOString();
      else if (typeof v === 'string' || v === null) out[(info as any).InternalName] = v;
      continue;
    }

    // Number/Currency
    if (isNumberField(info)) {
      const num = normalizeNumber(v);
      out[(info as any).InternalName] = num;
      continue;
    }

    // Boolean
    if (isBooleanField(info)) {
      const b = normalizeBoolean(v);
      out[(info as any).InternalName] = b;
      continue;
    }

    // Default primitives
    if (v === null || typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') {
      out[(info as any).InternalName] = v;
      continue;
    }
  }

  return out;
}


export async function savePmoItem(
  sp: SPFI,
  pmoListTitle: string,
  id: number,
  draft: any
): Promise<any> {
  const list = sp.web.lists.getByTitle(pmoListTitle);
  const fieldMap = await getFieldInfoMap(sp, pmoListTitle);
  const clean = buildCleanUpdatePayload(draft, fieldMap);

  await list.items.getById(id).update(clean);
  const saved = await list.items.getById(id).select('*')();
  return saved;
}

// =========================================================================
//   ==  הרשאות שדה ברמת PMO decisions לפי fieldPermission + קבוצות ==
//   ========================================================================= 

// טוען מפה: internalFieldName -> מערך תפקידי הרשאה (בחירות מ- WhoCanEdite) 
export async function loadFieldPermissionMap(
  sp: SPFI,
  fieldPermissionListTitle: string = FIELD_PERMISSION_LIST_TITLE
): Promise<Record<string, string[]>> {
  console.log("fieldPermissionListTitle ", fieldPermissionListTitle);
  const list = sp.web.lists.getByTitle(fieldPermissionListTitle);
  console.log("FP_COL_INTERNAL ", FP_COL_INTERNAL, "FP_COL_WHO_CAN_EDIT ", FP_COL_WHO_CAN_EDIT);
  const items = await list.items.select('Id', FP_COL_INTERNAL, FP_COL_WHO_CAN_EDIT)();
  console.log("✏️ who cab=n edite items: ", items);
  const map: Record<string, string[]> = {};
  for (let i = 0; i < items.length; i++) {
    const it = items[i] || {};
    const internal = it[FP_COL_INTERNAL];
    if (!internal) continue;

    let choices: string[] = [];
    const raw = it[FP_COL_WHO_CAN_EDIT];
    if (Array.isArray(raw)) {
      choices = raw as string[];
    } else if (raw && Array.isArray(raw.results)) {
      choices = (raw.results as string[]);
    } else if (typeof raw === 'string' && raw) {
      // Single choice as string
      choices = [raw];
    }

    if (!map[internal]) map[internal] = [];
    for (let j = 0; j < choices.length; j++) {
      const c = String(choices[j]);
      // מניעת כפילויות
      if (map[internal].indexOf(c) === -1) map[internal].push(c);
    }
  }
  return map;
}

// ✅ ייצוא בשם מהמקור
export async function fetchOrCreatePmoByIntegration(
  sp: SPFI,
  pmoListTitle: string,
  integrationItemId: number,
  integrationLookupDisplayOrInternal: string
): Promise<{ item: any; isNew: boolean }> {
  const list = sp.web.lists.getByTitle(pmoListTitle);

  // resolve lookup internal name
  const fields = await list.fields.select('Title','InternalName')();
  let lookupInternal = integrationLookupDisplayOrInternal;
  for (const f of fields as any[]) {
    const wantLc  = integrationLookupDisplayOrInternal.toLowerCase();
    if (String(f.Title || '').toLowerCase() === wantLc ||
        String(f.InternalName || '').toLowerCase() === wantLc) {
      lookupInternal = f.InternalName;
      break;
    }
  }

  const found = await list.items.filter(`${lookupInternal}Id eq ${integrationItemId}`).top(1)();
  if (found && found.length) return { item: found[0], isNew: false };

  const created = await list.items.add({ [`${lookupInternal}Id`]: integrationItemId });
  const item = await list.items.getById(created.data.Id).select('*')();
  return { item, isNew: true };
}


// שליפת כל המשתמשים ברשימת GeneralRoleDefinition לשלוש קבוצות /
export async function loadGeneralRoleUsers(
  sp: SPFI
): Promise<{
  FinancialAdvisor: string[]; // emails lower-case
  Lawyer: string[];
  PMOIntegrationTeam: string[];
}> {
  const list = sp.web.lists.getByTitle(GENERAL_ROLE_DEF_LIST_TITLE);
  // נניח שיש רק רשומה אחת עם עמודות PEOPLE מרובות; אם יש רבות – נאחד את כולן
  const rows = await list.items.select('Id', `${COL_FINANCIAL_ADVISOR}/EMail`, `${COL_LAWYER}/EMail`, `${COL_PMO_INTEGRATION_TEAM}/EMail`)
    .expand(COL_FINANCIAL_ADVISOR, COL_LAWYER, COL_PMO_INTEGRATION_TEAM)();

  const fin: string[] = [];
  const law: string[] = [];
  const pmo: string[] = [];

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || {};
    const addPeople = (arr: any, target: string[]) => {
      if (!arr) return;
      const src = Array.isArray(arr) ? arr : (arr.results || []);
      for (let j = 0; j < src.length; j++) {
        const p = src[j];
        const email = lc(p && (p.EMail || p.Email || p.UserPrincipalName));
        if (email && target.indexOf(email) === -1) target.push(email);
      }
    };
    addPeople(r[COL_FINANCIAL_ADVISOR], fin);
    addPeople(r[COL_LAWYER], law);
    addPeople(r[COL_PMO_INTEGRATION_TEAM], pmo);
  }

  return {
    FinancialAdvisor: fin,
    Lawyer: law,
    PMOIntegrationTeam: pmo,
  };
}

export async function loadTenderTeamUsersFromIntegration(
  sp: SPFI,
  integrationItemId: number
): Promise<string[]> {
  // שליפה של השדה כטקסט – ללא expand
  console.log("🫙🔪🎡🎊🎀🎁🎉✨🧨🎆🎈 in loadTenderTeamUsersFromIntegration");
  const it = await integrationList(sp).items
    .getById(integrationItemId)
    .select('Id', COL_PMO_TENDER_TEAM_FROM_CREATER)();
  console.log("it ", it);
  // אם השדה ריק – נחזיר ריק
  const raw = String(it?.[COL_PMO_TENDER_TEAM_FROM_CREATER] || '');
  console.log("🪂 COL_PMO_TENDER_TEAM_FROM_CREATER ", COL_PMO_TENDER_TEAM_FROM_CREATER, "raw ", raw);
  // מאפשר פסיקים/נקודה-פסיק/שורות מרובות
  const parts = raw.split(/[,;]|\r?\n/);
  console.log("🎈parts ", parts);
  const emails: string[] = [];
  for (let i = 0; i < parts.length; i++) {
    const e = parts[i].trim().toLowerCase();
    if (!e) continue;
    // ולידציה בסיסית של אימייל + מניעת כפילויות
    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e) && emails.indexOf(e) === -1) {
      emails.push(e);
    }
  }
  console.log("📧 emails ", emails); 
  return emails;
}


// החלטת עריכה לשדה ספציפי עבור משתמש נתון (email) /
export function canUserEditField(
  userEmailLc: string,
  internalName: string,
  fieldPermMap: Record<string, string[]>,
  roleUsers: {
    FinancialAdvisor: string[];
    Lawyer: string[];
    PMOIntegrationTeam: string[];
    PMOTenderTeam: string[]; // מגיע מ-Integration.FromCreater
  }
): boolean {
  const choices = fieldPermMap[internalName];
  if (!choices || !choices.length) return false; // ברירת מחדל: קריאה בלבד

  // בניית איחוד משתמשים לפי הבחירות
  const allowEmails: string[] = [];
  const pushDistinct = (src: string[]) => {
    for (let i = 0; i < src.length; i++) {
      const e = src[i];
      if (allowEmails.indexOf(e) === -1) allowEmails.push(e);
    }
  };

  for (let i = 0; i < choices.length; i++) {
    const choice = choices[i];
    const cLc = lc(choice);

    if (cLc.indexOf(lc('יועץ פיננסי')) > -1) pushDistinct(roleUsers.FinancialAdvisor);
    else if (cLc.indexOf(lc('עורך דין')) > -1) pushDistinct(roleUsers.Lawyer);
    else if (cLc.indexOf(lc('PMO צוות אינטגרציה')) > -1) pushDistinct(roleUsers.PMOIntegrationTeam);
    else if (cLc.indexOf(lc('PMO צוות מכרז')) > -1) pushDistinct(roleUsers.PMOTenderTeam);
  }

  return allowEmails.indexOf(userEmailLc) > -1;
}

*/