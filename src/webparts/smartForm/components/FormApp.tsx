//src\webparts\smartForm\components\FormApp.tsx


import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import {
  Stack, Pivot, PivotItem,
  PrimaryButton, DefaultButton,
  MessageBar, MessageBarType, 
  ComboBox, IComboBoxOption, Dropdown
} from '@fluentui/react';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { SPFI } from '@pnp/sp';
import { getSP } from './pnpjsConfig';
import { isSystemField } from '../shared/constants';
import { INTEGRATION_LIST_ID, INTEGRATION_PREVIEW_FIELD_INTERNAL } from '../shared/internalNames';
import {
  //getFieldMapsByTitle,
  //getFieldInfoMap,
  getFieldInfoMapById,
  fetchIntegrationItemByGuid,
  fetchOrCreatePmoByIntegration,
  savePmoItem,
  loadFieldPermissionMap,
  loadGeneralRoleUsers,
  loadTenderTeamUsersFromIntegration,
  canUserEditField, 
  FieldPermissionMap,
  getFieldMapsById
} from '../shared/data';
import stepsConfigJson from '../stepsConfig.json';
import '@pnp/sp/site-users/web';
import EditableFields from './EditableFields';
import '@pnp/sp/views';
import '@pnp/sp/lists';


// ===== עיצוב בסיסי (צבעים, כרטיסים וכו') =====
const PAGE_BG = 'linear-gradient(135deg, #f4f6fb 0%, #e7f2ff 40%, #f9fafb 100%)';
const CARD_BG = '#ffffff';
const CARD_SHADOW = '0 14px 40px rgba(15, 23, 42, 0.12)';
const CARD_RADIUS = 18;
const ACCENT = '#00498f';
const ACCENT_SOFT = '#e6f0ff';// ===== Reload Guard (per-user) =====
const RELOAD_GUARD_LIST_TITLE = "ReloadGuard";
const RELOAD_GUARD_USER_FIELD = "User";          // Person field
const RELOAD_GUARD_FLAG_FIELD = "HasReloadedOnce"; // Boolean field

function normalizePayloadForSpAdd(payload: any) {
  const out: any = {};
  for (const [k, v] of Object.entries(payload)) {
    // null/undefined נשארים
    if (v == null) { out[k] = v; continue; }

    // אם הגיע בצורה { results: [...] } → להמיר למערך [...]
    if (typeof v === "object" && !Array.isArray(v) && Array.isArray((v as any).results)) {
      out[k] = (v as any).results;
      continue;
    }

    out[k] = v;
  }
  return out;
}


function normalizeForSp(v: any) {
  if (v === undefined) return undefined; // לא שולחים בכלל
  if (v === null) return null;           // שולחים null
  if (Array.isArray(v)) return { results: v }; // MultiChoice/MultiText
  if (typeof v === "object" && v?.results) return v; // כבר בפורמט results
  return v;
}

export function buildIntegrationPayloadFromPmo(
  integrationItem: any, 
  pmoItem: any,
  pmoToIntegrationMap: Record<string, string>,
  extra: Record<string, any> = {}
) {
  const payload: any = {};

  for (const [k, v] of Object.entries(integrationItem ?? {})) {
    const nv = normalizeForSp(v);
    if (nv !== undefined) payload[k] = nv;
  }
  for (const [pmoInternal, integrationInternal] of Object.entries(pmoToIntegrationMap)) {
    const v = pmoItem?.[pmoInternal];
    const nv = normalizeForSp(v);
    if (nv !== undefined) payload[integrationInternal] = nv;
  }


  // תוספות/override (למשל TenderNumber לכל פריט חדש)
  for (const [k, v] of Object.entries(extra)) {
    const nv = normalizeForSp(v);
    if (nv !== undefined) payload[k] = nv;
  }


  return payload;
}
export async function splitTenderAndCreateIntegrationItems(params: {
  sp: any;
  integrationListId: string;      // e5e8eaea-...
  pmoItem: any;                   // הפריט של PMO (מלא)
  itegrationItem: any;
  tenderSourceInternalName: string; // למשל "DecisionAppliesToOtherWorksTende" או "TenderNumber"
  integrationTenderInternalName?: string; // ברירת מחדל "TenderNumber"
  pmoToIntegrationMap: Record<string, string>; // PMO_TO_INTEGRATION_FIELD_MAP
  // אם יש לכם עמודת lookup ב-Integration שמצביעה ל-PMO/Reference - תוסיפי כאן:
  linkFieldInternalName?: string; // למשל "NTA_x2019_s_x0020_reference"
  linkValue?: number;            // למשל pmoItem.IntegrationId או pmoItem.Id (מה שמתאים אצלכם)
}) {
  const {
    sp,
    integrationListId,
    pmoItem,
    itegrationItem,
    tenderSourceInternalName,
    integrationTenderInternalName = "TenderNumber",
    pmoToIntegrationMap,
    linkFieldInternalName,
    linkValue,
  } = params;

  const raw = String(pmoItem?.[tenderSourceInternalName] ?? "").trim();
  const parts = raw
    .split(",")
    .map(s => s.trim())
    .filter(Boolean);

  if (parts.length === 0) return;

  const list = sp.web.lists.getById(integrationListId);

  for (const tender of parts) {
    const extra: any = { [integrationTenderInternalName]: tender };

    // אם יש lookup/Reference חובה ברשימת Integration
    if (linkFieldInternalName && linkValue != null) {
      extra[linkFieldInternalName] = linkValue;
    }

    const payload = buildIntegrationPayloadFromPmo(itegrationItem, pmoItem, pmoToIntegrationMap, extra);
    console.log("🏔️ integration addPayload", payload);

    const fixed = normalizePayloadForSpAdd(payload);
    console.log("✅ fixed payload", fixed);
    await list.items.add(fixed);
  }
}


export type FieldInfoLike = {
  InternalName: string;
  TypeAsString?: string;
  Hidden?: boolean;
  ReadOnlyField?: boolean;
  Sealed?: boolean;
};

function splitCommaList(raw: string): string[] {
  return raw
    .split(",")
    .map(s => s.trim())
    .filter(Boolean);
}

function isEmptyValue(v: any): boolean {
  if (v === null || v === undefined) return true;
  if (typeof v === "string") return v.trim() === "";
  if (Array.isArray(v)) return v.length === 0;
  return false;
}

function isSystemFieldLocal(internal: string): boolean {
  const s = internal.toLowerCase();
  return (
    s === "id" ||
    s === "guid" ||
    s === "attachments" ||
    s === "author" || s === "authorid" ||
    s === "editor" || s === "editorid" ||
    s === "created" ||
    s === "modified" ||
    s === "contenttypeid" ||
    s.startsWith("odata") ||
    s.startsWith("odata__") ||
    s.startsWith("_") ||
    s.startsWith("filesystemobjecttype") ||
    s.startsWith("complianceassetid") ||
    s.startsWith("serverredirected") ||
    s.startsWith("owshiddenversion") ||
    s.startsWith("uiversion") ||
    s.startsWith("odata__uiversionstring") ||
    s.startsWith("odata__colortag")
  );
}

function pickValue(primary: any, secondary: any, internal: string) {
  const a = primary?.[internal];
  if (!isEmptyValue(a)) return a;
  const b = secondary?.[internal];
  if (!isEmptyValue(b)) return b;
  return undefined;
}

/**
 * הופך ערכים לפורמט ש-SharePoint REST מצפה לו במקרים נפוצים.
 * (הכי חשוב: MultiChoice => { results: [...] })
 */
function coerceValueForSp(typeAsString: string, val: any) {
  const t = (typeAsString || "").toLowerCase();

  // תאריך
  if (t.includes("date") || t.includes("datetime")) {
    if (val instanceof Date) return val.toISOString();
    return val; // אם זה כבר ISO string
  }

  // MultiChoice / MultiText / Multi (בדרך כלל SharePoint רוצה results)
  if (Array.isArray(val)) {
    // בהרבה מקרים זה MultiChoice
    return { results: val };
  }

  // URL/Hyperlink כבר את מנרמלת לפני save; אם לא - נשאיר כמו שהוא
  return val;
}

function buildMergedClonePayload(
  primary: any,
  secondary: any,
  fieldInfoMap: Record<string, FieldInfoLike>,
  overrides: Record<string, any>
) {
  const payload: any = {};

  for (const internal of Object.keys(fieldInfoMap)) {
    const info = fieldInfoMap[internal];
    if (!info) continue;

    if (info.Hidden || info.ReadOnlyField || info.Sealed) continue;
    if (isSystemFieldLocal(internal)) continue;

    // לוקחים ערך מהראשון, ואם ריק אז מהשני
    const v = pickValue(primary, secondary, internal);
    if (v === undefined) continue;

    payload[internal] = coerceValueForSp(info.TypeAsString || "", v);
  }

  // overrides (למשל TenderNumber = X)
  for (const k of Object.keys(overrides)) {
    payload[k] = overrides[k];
  }

  return payload;
}

export async function splitTenderAndCreateItemsFromTwoSources(params: {
  sp: any;
  listId: string;                    // הרשימה הראשונה 
  primaryItem: any;                  // item A (למשל PMO)
  secondaryItem: any;                // item B (למשל Integration snapshot)
  fieldInfoMap: Record<string, FieldInfoLike>;
  tenderFieldInternalName?: string;  // ברירת מחדל TenderNumber

}) {
  const {
    sp,
    listId,
    primaryItem,
    secondaryItem,
    fieldInfoMap,
    tenderFieldInternalName = "DecisionAppliesToOtherWorksTende",
  } = params;
  console.log("🏔️🏔️🏔️ primaryItem ", primaryItem);
  console.log("🏔️🏔️🏔️ secondaryItem ", secondaryItem);
  const raw =
    String(primaryItem?.[tenderFieldInternalName] ?? secondaryItem?.[tenderFieldInternalName] ?? "").trim();
  console.log("🏔️🏔️🏔️raw ", raw);
  const parts = splitCommaList(raw);
  if (parts.length <= 1) return;
  console.log("🏔️🏔️🏔️parts ", parts );
  const list = sp.web.lists.getById(listId);

  

  // ליצור פריטים חדשים עבור שאר החלקים
  for (const x of parts.slice(0)) {
    const addPayload = buildMergedClonePayload(
      primaryItem,
      secondaryItem,
      fieldInfoMap,
      { ['TenderNumber']: x }
    );
    console.log("🏔️🏔️🏔️addPayload ", addPayload);

    await list.items.add(addPayload);
  }
}


// מחזיר: ["M3"] או ["M2","M3"] או ["ALL"] או שילובים
export async function myOLM(
  sp: SPFI,
  generalRoleDefinitionListId: string,
  email: string
): Promise<Array<"M1" | "M2" | "M3" | "ALL">> {

  console.log("🎍 in my OLM");

  const userEmail = String(email || "").trim().toLowerCase();
  if (!userEmail) return [];
  console.log("🎍 hase email ");
  const ROLE_COLUMNS: Record<"M1" | "M2" | "M3", string[]> = {
    M1: ["M1TenderTeam", "LawyerM1TenderTeam"],
    M2: ["M2TenderTeam", "LawyerM2TenderTeam"],
    M3: ["M3TenderTeam", "LawyerM3TenderTeam"],
  };

  const ALL_COLUMNS = ["IntegrationTeamLawyer", "IntegrationTeam", "FinancialAdvisorIntegrationTeam"];

  

   console.log("🎍 1 ");
  //const select = ["Id", ...peopleCols].join(",");
  console.log("🎍 2 ");
  const list = sp.web.lists.getById(generalRoleDefinitionListId);
console.log("🎍 3 ");
  // ⬅️ מביאים את כל הפריטים (לא רק הראשון)
  const rows: any[] = await list.items
    .select()();
console.log("🎍 4 , rows ", rows);
  if (!rows?.length) return [];
console.log("🎍 5 ");
  const roles = new Set<"M1" | "M2" | "M3" | "ALL">();
console.log("🎍 6 ");
  const emailInTextField = (fieldVal: any): boolean => {
    if (!fieldVal) return false;
    return String(fieldVal).trim().toLowerCase() === userEmail;
  };
  console.log("🎍 7 ");

  for (const row of rows) {

    // ALL
    for (const col of ALL_COLUMNS) {
      if (emailInTextField(row[col])) {
        roles.add("ALL");
      }
    }

    // M1 / M2 / M3
    (Object.keys(ROLE_COLUMNS) as Array<"M1" | "M2" | "M3">).forEach(role => {
      for (const col of ROLE_COLUMNS[role]) {
        if (emailInTextField(row[col])) {
          console.log("emailInTextField(row[col]  ", row[col])
          roles.add(role);
        }
      }
    });
  }
console.log("🎍 8 ");
  const result = Array.from(roles);
  console.log("🎍 roles:", result);

  return result;
}


async function ensureReloadGuardList(sp: any) {
  let list;
  console.log("🕐 Starting ensureReloadGuardList");
  
  try {
    // ⚠️ CRITICAL: Ensure sp.web is actually initialized
    if (!sp?.web) {
      throw new Error("SP context not initialized");
    }
    
    // Test if list exists
    list = sp.web.lists.getByTitle(RELOAD_GUARD_LIST_TITLE);
    await list.select("Id")(); // Lightweight check
    console.log("✅ List exists");
    
  } catch (err: any) {
    console.log("📝 List doesn't exist, creating...", err.message);
    
    // Create the list
    const res = await sp.web.lists.add(RELOAD_GUARD_LIST_TITLE, "", 100, false);
    list = res.list;
  }
  
  // Ensure fields exist...
  // (rest of your field creation code)
  
  return list;
}

async function getOrCreateReloadGuardItem(sp: any, list: any) {
  const me = await sp.web.currentUser();
  const items = await list.items
    .filter(`${RELOAD_GUARD_USER_FIELD}/Id eq ${me.Id}`)
    .select("Id", RELOAD_GUARD_FLAG_FIELD, `${RELOAD_GUARD_USER_FIELD}/Id`)
    .expand(RELOAD_GUARD_USER_FIELD)();

  if (items.length > 0) return items[0];

  const addRes = await list.items.add({
    [`${RELOAD_GUARD_USER_FIELD}Id`]: me.Id,
    [RELOAD_GUARD_FLAG_FIELD]: false
  });

  return { Id: addRes.data.Id, [RELOAD_GUARD_FLAG_FIELD]: false };
}

async function markReloadedAndReload(sp: any) {
  const list = await ensureReloadGuardList(sp);
  const item = await getOrCreateReloadGuardItem(sp, list);

  const alreadyReloaded = !!item?.[RELOAD_GUARD_FLAG_FIELD];
  if (!alreadyReloaded) {
    await list.items.getById(item.Id).update({
      [RELOAD_GUARD_FLAG_FIELD]: true
    });
    window.location.reload();
    return true; // did reload
  }
  return false; // no reload
}

async function resetReloadGuard(sp: any) {
  try {
    const list = await ensureReloadGuardList(sp);
    const me = await sp.web.currentUser();

    const items = await list.items
      .filter(`${RELOAD_GUARD_USER_FIELD}/Id eq ${me.Id}`)
      .select("Id", RELOAD_GUARD_FLAG_FIELD, `${RELOAD_GUARD_USER_FIELD}/Id`)
      .expand(RELOAD_GUARD_USER_FIELD)();

    if (items.length > 0) {
      await list.items.getById(items[0].Id).update({
        [RELOAD_GUARD_FLAG_FIELD]: false
      });
    }
  } catch {
    // לא נורא אם נכשל — זה רק מנגנון עזר
  }
}


function htmlToPlainText(html?: string): string {
  if (!html) return '';
  let s = String(html);
  s = s
    .replace(/<\s*li[^>]*>/gi, '• ')
    .replace(/<\s*br\s*\/?>/gi, '\n')
    .replace(/<\/\s*p\s*>/gi, '\n')
    .replace(/<\/\s*li\s*>/gi, '\n');
  s = s.replace(/<[^>]+>/g, '');
  const ta = document.createElement('textarea');
  ta.innerHTML = s;
  return ta.value;
}

function normalizeStepsConfigToInternal(
  steps: Record<string, string[]>,
  titleToInternal: Record<string, string>
): Record<string, string[]> {
  const out: Record<string, string[]> = {};
  const stepNames = Object.keys(steps || {});
  for (let s = 0; s < stepNames.length; s++) {
    const step = stepNames[s];
    const arr = steps[step] || [];
    const internals: string[] = [];
    const seen: Record<string, boolean> = {};
    for (let i = 0; i < arr.length; i++) {
      const name = arr[i];
      if (!name) continue;
      const internal = titleToInternal[name] || name;
      if (internal && !seen[internal]) { internals.push(internal); seen[internal] = true; }
    }
    out[step] = internals;
  }
  return out;
}

function getPhaseViewKeyFromTenderPhase(tenderPhaseRaw?: string): string {
  const s = String(tenderPhaseRaw || '').trim().toLowerCase();
  if (!s) return '';

  // אם מתחיל ב־Phase 1 / Phase 2 / Phase 3 – ניקח את ה־prefix
  if (s.indexOf('Phase 1')!= -1)  return 'Phase 1';
  if (s.indexOf('Phase 2')!= -1) return 'Phase 2';
  if (s.indexOf('Phase 3')!= -1) return 'Phase 3';

  // ברירת מחדל – אין סינון לפי תצוגה
  return '';
}




export interface FormAppProps {
  context: WebPartContext;
  pmoListTitle: string;
  pmoIntegrationLookupName?: string;
  stepsConfig?: Record<string, string[]>;
  isEditMode?: boolean; 
}

const FormApp: React.FC<FormAppProps> = ({
  context,
  pmoListTitle,
  pmoIntegrationLookupName = 'Integration',
  stepsConfig = stepsConfigJson,
  isEditMode = false
}) => {
  const PMO_LIST_ID = "e5e8eaea-16db-49d3-ad7c-62f5a2bdd97a"; 
  const sp: SPFI = useMemo(() => getSP(context), [context]);


  const [me, setMe] = useState<{ Email?: string; Title?: string; LoginName?: string } | null>(null);

  const [integrationChoices, setIntegrationChoices] = useState<IComboBoxOption[]>([]);
  const [integrationId, setIntegrationId] = useState<number | null>(null);
  // const [comboOpen, setComboOpen] = useState<boolean>(false);
  const [olmFilter, setOlmFilter] = useState<string>('ALL');



  // Integration (תצוגה בלבד) + תוויות + מטא
  const [integrationItem, setIntegrationItem] = useState<any>(null);
  const [integrationFieldInfoMap, setIntegrationFieldInfoMap] = useState<Record<string, any>>({});
  const [integrationLabels, setIntegrationLabels] = useState<Record<string, string>>({});

    // סדר שדות לתצוגה לפי View של רשימת Integration
  const [integrationViewFieldOrder, setIntegrationViewFieldOrder] = useState<string[]>([]);
  //const [, setIntegrationViewName] = useState<string>('');
  const [integrationSearch, setIntegrationSearch] = useState<string>('');


  // PMO (עריכה) + תוויות + מטא
  const [pmoItem, setPmoItem] = useState<any>(null);
  const [pmoDraft, setPmoDraft] = useState<any>({});

  // placeholder לשדות ספציפיים כשיש Accept
  const placeholders = React.useMemo<Partial<Record<string, string>>>(() => {
    const decision = String(pmoDraft?.DecisionRegardingProposedChange || '').trim();
    if (decision === 'Accept') {
      const msg = 'Enter final wording for publication here';
      return {
        RevisedWordingFinalForPublicatio: msg,
        dog: msg
      };
    }
    return {};
  }, [pmoDraft?.DecisionRegardingProposedChange]);

  const [pmoLabels, setPmoLabels] = useState<Record<string, string>>({});
  const [pmoFieldInfoMap, setPmoFieldInfoMap] = useState<Record<string, any>>({});

  // שלבים
  const [stepsInternal, setStepsInternal] = useState<Record<string, string[]>>({});
  const [activeStep, setActiveStep] = useState<string>('');

    // שדות לפי תצוגת SharePoint (View) לפי TenderPhase
  const [viewFieldOrder, setViewFieldOrder] = useState<string[]>([]);
  const [currentViewName, setCurrentViewName] = useState<string>('');


  // הרשאות שדה
  const [fieldPermMap, setFieldPermMap] = useState<FieldPermissionMap>({});

  const [myRoles, setMyRoles] = React.useState<
    Array<"M1" | "M2" | "M3" | "ALL">
  >([]);

  const [roleUsers, setRoleUsers] = useState<{
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
  }>({
    IntegrationTeam: [],
    IntegrationTeamLawyer: [],
    FinancialAdvisorIntegrationTeam: [],
    M1TenderTeam: [],
    LawyerM1TenderTeam: [],
    M2TenderTeam: [],
    LawyerM2TenderTeam: [],
    M3TenderTeam: [],
    LawyerM3TenderTeam: [],
    FinancialAdvisor: [],
    Lawyer: [],
    PMOIntegrationTeam: [],
    PMOTenderTeam: []
  });

  // UI
  const [busy, setBusy] = useState<boolean>(false);
  const [msg, setMsg] = useState<{ type: MessageBarType; text: string } | null>(null);

  // ----- שדות חובה בטופס PMO decisions -----
  const REQUIRED_FIELDS: string[] = [
    'DecisionRegardingProposedChange',
    'DecisionAppliesToOtherWorksTende',
    'RevisionIncludesChangeInTenderDo',
    'DecisionDate'
  ];

  const TENDER_TEAM_FIELDS: string[] = [
  "StatusOfRFCresponseOrTcRFC",//
  "TenderCommitteeApprovalDate",//
  "Addendum",//
  "RFCresponseAsPublishedToBeFilled",//
  "RFCorTcRFCasPublishedByNTaToBeFi",//
  "RFCResponseLetterNo",//
  "RevisedWordingFinalForPublicatio",//
  "IntegrationTeamDecisionImplement",//
  'DueDateCalculated',
  'ActualDate'
];

  // 🆕 מפת תוויות קשיחה – fallback במקרה שאין internalToTitle מה-SharePoint
const PMO_LABEL_OVERRIDES: Record<string, string> = {
  Integration:'NTA reference#',
  DecisionRegardingProposedChange:'Decision Regarding Proposed Change',
  DecisionRegardingProposedChangeC:'Decision Regarding Proposed Change - comments', 
  DecisionAppliesToOtherWorksTende:'Decision applies to other Works Tenders?',
  IntegrationTeamDecisionImplement:'Decision Implementation Status',
  RevisionIncludesChangeInTenderDo: 'Revision includes change in Tender Documents? (Y/N) If Y, Addendum required', 
  RFCResponseLetterNo:'RFC response Letter no.',
  RFCresponseAsPublishedToBeFilled:'RFC response as published  (To be filled in after publication)',
  Addendum:'Addendum #',
  TenderCommitteeApprovalDate:'Tender Committee Approval Date',
  StatusOfRFCresponseOrTcRFC:'Status Of RFC response Or Tc RFC',
  DecisionDate:'Decision Date',
  RevisedWordingFinalForPublicatio:'Revised Wording - final for publication',
  RFCorTcRFCasPublishedByNTaToBeFi:'RFC / TC RFC as published by NTA (To be filled in after publication)',
  DueDateCalculated:'Due date (calculated)',
  ActualDate:'Actual date',
  DateForIntegrationTeamDecisionIm:'Date For Integration Team Decision Implementation Status',
  INTEGRATIONTEAMSTATUS:'INTEGRATION TEAM STATUS',
  RFCorTcRFCasPublishedByNTaNew: 'RFCorTcRFCasPublishedByNTaNew',
  sentProtocol:'sentProtocol',
};



  // מפה מהיר מ־internalName -> חובה?
  const requiredMap: Record<string, boolean> = useMemo(() => {
    const m: Record<string, boolean> = {};
    for (let i = 0; i < REQUIRED_FIELDS.length; i++) {
      m[REQUIRED_FIELDS[i]] = true;
    }
    return m;
  }, []);

  // שגיאות ולידציה לשדות (שדה -> טקסט שגיאה)
  const [validationErrors, setValidationErrors] = useState<Record<string, string>>({});
  // ==== Sync PMO -> INTEGRATION ====

// מפה בין שמות השדות ב־PMO לבין השדות המקבילים ברשימת INTEGRATION
// 🟢 אם השמות זהים בשתי הרשימות – פשוט תשאירי את אותו שם גם מימין וגם משמאל.
// 🟢 תעדכני כאן את כל השדות שאת רוצה לסנכרן.
const PMO_TO_INTEGRATION_FIELD_MAP: Record<string, string> = {
  // PMO internal name      : INTEGRATION internal name
  DecisionRegardingProposedChange: 'DecisionRegardingProposedChange',
  DecisionRegardingProposedChangeC: 'DecisionRegardingProposedChange_',
  DecisionAppliesToOtherWorksTende: 'DecisionappliestootherWorksTende',
  IntegrationTeamDecisionImplement: 'IntegrationTeamDecision',
  RevisionIncludesChangeInTenderDo: 'RevisionincludeschangeinTenderDo',
  RFCResponseLetterNo: 'RFCresponseLetterno',
  RFCresponseAsPublishedToBeFilled: 'RFCresponseaspublished',
  Addendum: 'Addendum',
  TenderCommitteeApprovalDate: 'TenderCommitteeapprovaldate',
  StatusOfRFCresponseOrTcRFC: 'StatusofRFCresponse_x002f_TCRFC',
  DecisionDate: 'Decisiondate',
  RevisedWordingFinalForPublicatio: 'RevisedWording_x002d_finalforpub',
  dog:'RFCorTcRFCasPublishedByNTaNew', //'RFC_x002f_TCRFCaspublishedbyNTA_',
  DueDateCalculated: 'Duedate',
  ActualDate: 'Actualdate',
  // וכו' – תוסיפי/תמחקי לפי הצורך
};


async function syncPmoToIntegration(
  sp: SPFI,
  integrationId: number,
  pmoDraft: any
): Promise<void> {
  if (!integrationId) return;

  const updatePayload: any = {};

    for (const pmoField in PMO_TO_INTEGRATION_FIELD_MAP) {
      if (!Object.prototype.hasOwnProperty.call(PMO_TO_INTEGRATION_FIELD_MAP, pmoField)) continue;

      const integrationField = PMO_TO_INTEGRATION_FIELD_MAP[pmoField];
      const val = pmoDraft ? pmoDraft[pmoField] : undefined;

      if (typeof val === 'undefined') continue;

      let outVal = val;


      updatePayload[integrationField] = outVal;
    }


  if (Object.keys(updatePayload).length === 0) {
    console.log('syncPmoToIntegration: nothing to update');
    return;
  }

  console.log('syncPmoToIntegration →', updatePayload);

  await sp.web.lists
    .getById(INTEGRATION_LIST_ID)
    .items.getById(integrationId)
    .update(updatePayload);
}

  const filteredIntegrationChoices: IComboBoxOption[] = useMemo(() => {
    console.log("🦝🦝🦝🦝🦝🦝🦝🦝 filteredIntegrationChoices")
    //const q = integrationSearch.trim().toLowerCase();
    const q = (integrationSearch || '').trim().toLowerCase();

    const arr = integrationChoices as IComboBoxOption[];

    return arr.filter(opt => {
      const textMatch =
        !q || String(opt.text || '').toLowerCase().indexOf(q) !== -1;

      const olm = String((opt as any).data?.olm || '').toUpperCase();

      const olmMatch =
        olmFilter === 'ALL' || !olmFilter
          ? true
          : olm === olmFilter;

      return textMatch && olmMatch;
    });
  }, [integrationSearch, integrationChoices, olmFilter]);



    const loadIntegrationChoices = async () => {
      console.log("loadIntegrationChoices ")

      const items: any[] = await sp.web.lists
        .getById(INTEGRATION_LIST_ID)
        .items.select(
          'Id',
          INTEGRATION_PREVIEW_FIELD_INTERNAL,
          'NTA_x2019_s_x0020_reference',
          'OriginatingLineManager' // 👈 מוסיפים את השדה לפעולת ה־select
        )
        .orderBy('Id', false)
        .top(200)();
      console.log("loadIntegrationChoices items ", items);

      const myRolesUpper = myRoles.map(r => r.toUpperCase());
      const showAll = myRolesUpper.includes('ALL');
      const opts: IComboBoxOption[] = [];

      for (let i = 0; i < items.length; i++) {
        const it = items[i];
        const id = it.Id as number;
        const raw = it[INTEGRATION_PREVIEW_FIELD_INTERNAL];
        const plain = htmlToPlainText(raw);
        const oneLine = plain.replace(/\r?\n/g, ' ').trim();
        const preview = oneLine.length > 30 ? oneLine.substring(0, 30) + '…' : oneLine;
        //const ntaRef = String(it.NTA_x2019_s_x0020_reference || '').trim();

        const olm = String(it.OriginatingLineManager || '').trim().toUpperCase(); // 👈 נחלץ את ה־OLM

        const text = preview ? `${id} — ${preview}` : String(id);
        
        if (!showAll && (!olm || !myRolesUpper.includes(olm))) {
          continue;
        }

        opts.push({
          key: id,
          text,
          // נשמור את ה-OLM בתוך data כדי שנוכל לסנן אחר כך
          data: { olm }
        });
      }
      console.log("🥗 opts ", opts);
      setIntegrationChoices(opts);
    };

    useEffect(() => {
      if (!sp) return;
      if (!myRoles.length) return; // ⬅️ זה החלק הקריטי

      loadIntegrationChoices();
    }, [sp, myRoles]);



  const loadIntegrationMeta = async () => {
    const fields = await sp.web.lists
      .getById(INTEGRATION_LIST_ID)
      .fields
      .select('Title', 'InternalName', 'Hidden', 'ReadOnlyField', 'TypeAsString')();

    const labels: Record<string, string> = {};
    const meta: Record<string, any> = {};

    for (let i = 0; i < fields.length; i++) {
      const f: any = fields[i];
      const internal = f.InternalName;
      labels[internal] = f.Title || internal;
      meta[internal] = {
        InternalName: internal,
        Title: f.Title,
        Hidden: !!f.Hidden,
        ReadOnlyField: !!f.ReadOnlyField,
        TypeAsString: f.TypeAsString
      };
    }
    setIntegrationLabels(labels);
    setIntegrationFieldInfoMap(meta);
  };

  const loadIntegrationLabels = async () => {
    const fields = await sp.web.lists.getById(INTEGRATION_LIST_ID).fields.select('Title', 'InternalName')();
    const map: Record<string, string> = {};
    for (let i = 0; i < fields.length; i++) {
      const f: any = fields[i];
      map[f.InternalName] = f.Title || f.InternalName;
    }
    setIntegrationLabels(map);
  };

  const getOriginatingLineManager = (): string => {
  const raw = String(integrationItem?.OriginatingLineManager || '').trim().toUpperCase();
  if (raw === 'M1' || raw === 'M2' || raw === 'M3') {
    return raw;
  }
  return '';
};

const getIntegrationViewNameForTenderPhase = (rawPhase: string): string => {
  const phase = (rawPhase || '').toLowerCase().trim();

  // כאן תתאימי לשמות התצוגות אצלך ברשימת Integration
  if (phase.indexOf('phase 1')===0) {
    return 'Phase 1';   // ← שם ה-View ברשימת INTEGRATION
  }
  if (phase.indexOf('phase 2')===0) {
    return 'Phase 2';
  }
  if (phase.indexOf('phase 3')===0) {
    return 'Phase 3';
  }

  // ברירת מחדל – למשל "All Items"
  return 'All Items';
};

// 🆕 טעינת סדר שדות לתצוגה מתוך View של רשימת Integration
const loadIntegrationViewFieldOrderForPhase = async (tenderPhaseRaw: string) => {
  console.log("🧩🧩🧩");
  try {
    const viewName = getIntegrationViewNameForTenderPhase(tenderPhaseRaw);

    // לוקחים את ה-View מרשימת INTEGRATION לפי ה-ID
    const vf: any = await sp.web.lists
      .getById(INTEGRATION_LIST_ID)
      .views.getByTitle(viewName)
       .fields();

    const internalNames: string[] = (vf?.Items || vf) as string[];

    console.log('🧩 Integration view fields for', viewName, internalNames);
    setIntegrationViewFieldOrder(internalNames || []);
  } catch (e) {
    console.error('Error loading Integration view fields for phase', e);
    setIntegrationViewFieldOrder([]); // במקרה של שגיאה – נציג הכל
  }
};


  // מחזיר שם תצוגה (View) לפי הערך של TenderPhase ברשימת Integration
  const getViewNameForTenderPhase = (rawPhase: string): string => {
    const phase = (rawPhase || '').toLowerCase();

    if (phase.indexOf('Phase 1') === 0) {
      return 'Phase 1';      // שם ה-View ברשימת PMO
    }
    if (phase.indexOf('Phase 2') === 0) {
      return 'Phase 2';      // שם ה-View ברשימת PMO
    }
    if (phase.indexOf('Phase 3') === 0) {
      return 'Phase 3';      // שם ה-View ברשימת PMO
    }
    console.log("");
    // ברירת מחדל – למשל All Items או מה שתרצי
    return 'כל הפריטים';
  };


    // טוען את רשימת העמודות מה-View המתאים (לפי TenderPhase)
  const loadViewFieldOrderForPhase = async (tenderPhaseRaw: string) => {
    console.log("🫏🐴🐎 loadViewFieldOrderForPhase");
    try {
      const viewName = getViewNameForTenderPhase(tenderPhaseRaw);
      console.log("🫏🐴🐎 1");
      setCurrentViewName(viewName);
      console.log("🫏🐴🐎 2");
      console.log("currentViewName ", currentViewName, " | ", viewName);
      console.log("🫏🐴🐎 2.5");
      /* לוקחים את ה-View מהרשימה של ה-PMO
      const viewFieldsResp: any = await sp.web.lists
        .getById(PMO_LIST_ID) 
        .views.getByTitle(viewName)
        .select('ViewFields')();
      */// זה האובייקט שרשמת ב-log

      // בקובץ data.ts או קובץ shared אחר
      
      const fields = await sp.web.lists
      .getById(PMO_LIST_ID)
      .views
      .getById("62CFFAB1-507D-43B5-A87F-4F4CCD64BAD2")
      .fields();
      console.log("🫏🐴🐎 3");

      type ViewFieldsResponse = {
        Items: string[];
        SchemaXml: string;
      };
      console.log("🫏🐴🐎 4");
      // ב-PnP בדרך כלל השמות הפנימיים נמצאים ב-Items או במערך עצמו
      //const internalNames: string[] =(viewFieldsResp?.ViewFields?.Items as string[]) || [];
      const resp = fields as ViewFieldsResponse;
      console.log("🫏🐴🐎 5");
      const internalNames: string[] = (resp.Items ) ||[];
      console.log("🫏🐴🐎 6");
      console.log("internalNames ", internalNames);
      setViewFieldOrder(internalNames || []);

      // אם את רוצה שכל טאב ב-Pivot יהיה בעצם "ה-View עצמו":
      setStepsInternal({
        [viewName]: internalNames || []
      });
      setActiveStep(viewName);
    } catch (e) {
      console.log("🫏🐴🐎 catch");
      console.error('🫏🐴🐎 Error loading view fields for phase', e);
      // במקרה של שגיאה – נשאיר את stepsInternal כמו שהיה, כדי שלא ישבר
    }
  };



  //🎀
  // האם להציג את RFCResponseLetterNo? רק אם RevisionIncludesChangeInTenderDo = 'Y'
  const showRFCResponseLetterNo = React.useMemo(() => {
    const v = String(pmoDraft?.RevisionIncludesChangeInTenderDo ?? '').trim().toUpperCase();
    return v === 'Y';
  }, [pmoDraft?.RevisionIncludesChangeInTenderDo]);

  React.useEffect(() => {
    console.log("🎍 in use effect trying to find your OLM");
    if (!sp || !me?.Email) return;
    console.log("🎍 !sp || !me?.Email");
    let cancelled = false;

    (async () => {
      const roles = await myOLM(
        sp,
        "b321608b-6405-4ad0-8676-11c7350fa7a4",
        String(me.Email)
      );
      

      if (!cancelled) {
        setMyRoles(roles);
      }
    })();
    
    return () => {
      cancelled = true;
    };
  }, [sp, me?.Email]);


  // אופציונלי: אם השדה מוסתר – ננקה את הערך כדי שלא יישמר בטעות
  React.useEffect(() => {
    if (!showRFCResponseLetterNo && pmoDraft?.RFCResponseLetterNo) {
      //setPmoDraft((prev: any) => ({ ...prev, RFCResponseLetterNo: null }));
    }
  }, [showRFCResponseLetterNo, pmoDraft?.RFCResponseLetterNo]);
  //🎀

  const loadFormForIntegration = async () => {

    console.log("🔮0");
    if (!integrationId) return;
    // Integration
    console.log("🔮1");
    const integ = await fetchIntegrationItemByGuid(sp, integrationId);
    console.log("🔮2");
    setIntegrationItem(integ);
    console.log("🔮3");

    // 🔹 אחרי שקיבלנו את פריט ה-Integration – נטען שדות לפי View
    const tenderPhaseRaw = String(integ?.TenderPhase || '');
    console.log("4🔮");
    await loadViewFieldOrderForPhase(tenderPhaseRaw);
    console.log("5🔮");
    await loadIntegrationViewFieldOrderForPhase(tenderPhaseRaw);
    console.log("6🔮");

    // PMO לפי lookup ל-Integration
    const { item: pmoFound, isNew } = await fetchOrCreatePmoByIntegration(
      sp, pmoListTitle, integrationId, pmoIntegrationLookupName
    );
    console.log("7🔮");
    setPmoItem(pmoFound);
    console.log("8🔮");
    setPmoDraft(pmoFound);
    console.log("9🔮");
    // טעינת הרשאות לפי fieldPermission + קבוצות גלובליות + צוות מכרז מתוך ה-Integration
    const [permMap, general, tenderTeamUsers] = await Promise.all([
      loadFieldPermissionMap(sp),
      loadGeneralRoleUsers(sp),
      null//loadTenderTeamUsersFromIntegration(sp, integrationId!)
    ]);
    console.log("10🔮");

    setFieldPermMap(permMap);
    console.log("11🔮");
    setRoleUsers({
      IntegrationTeam: general.IntegrationTeam,
      IntegrationTeamLawyer:  general.IntegrationTeamLawyer,
      FinancialAdvisorIntegrationTeam:  general.FinancialAdvisorIntegrationTeam,
      M1TenderTeam:  general.M1TenderTeam,
      LawyerM1TenderTeam:  general.LawyerM1TenderTeam,
      M2TenderTeam:  general.M2TenderTeam,
      LawyerM2TenderTeam:  general.LawyerM2TenderTeam,
      M3TenderTeam:  general.M3TenderTeam,
      LawyerM3TenderTeam:  general.LawyerM3TenderTeam,
      FinancialAdvisor: general.FinancialAdvisor,
      Lawyer: general.Lawyer,
      PMOIntegrationTeam: general.PMOIntegrationTeam,
      PMOTenderTeam: tenderTeamUsers?tenderTeamUsers:[""],   // ← חשוב: כאן נכנסת הרשאת "PMO צוות מכרז"
    });
    console.log("12🔮");

    if (isNew) setMsg({ type: MessageBarType.success, text: 'A new PMO item linked to Integration has been created.' });
    else setMsg(null);
    // Labels + Meta ל-PMO
    console.log("13🔮");
    //const pmoMaps = await getFieldMapsByTitle(sp, pmoListTitle);
    const pmoMaps = await getFieldMapsById(sp, PMO_LIST_ID);
    console.log("14🔮");
    setPmoLabels(pmoMaps.internalToTitle);
    console.log("15🔮");
    //const fieldMap = await getFieldInfoMap(sp, pmoListTitle);
    const fieldMap = await getFieldInfoMapById(sp, PMO_LIST_ID);
    console.log("16🔮");
    setPmoFieldInfoMap(fieldMap);
    console.log("17🔮");
    // steps → internal
    const normalized = normalizeStepsConfigToInternal(stepsConfig, pmoMaps.titleToInternal);
    console.log("18🔮");
    setStepsInternal(normalized);
    console.log("19🔮");
    const firstStep = Object.keys(normalized)[0] || '';
    console.log("20🔮");
    setActiveStep(firstStep);
    console.log("21🔮");
    // PMO צוות מכרז מתוך פריט ה-Integration
    const tenderTeam = await loadTenderTeamUsersFromIntegration(sp, integrationId);
    console.log("22🔮");
    setRoleUsers(prev => ({ ...prev, PMOTenderTeam: tenderTeam }));
    console.log("23🔮");
  };






  useEffect(() => {
    (async () => {
      try {
        const user = await sp.web.currentUser.select('Email', 'Title', 'LoginName')();
        setMe(user);
      } catch {
        setMe(null);
      }
    })();
  }, [sp]);

  const AUTO_RELOAD_KEY = 'SmartForm_AutoReloadOnce';
 
  useEffect(() => {
  (async () => {
    try {
      setBusy(true);
      
      // ✅ WAIT for SP to be fully initialized
      await sp.web.select("Title")(); // Simple call to ensure context is ready
      
      await Promise.all([
        loadIntegrationChoices(),
        loadIntegrationMeta(),
        loadIntegrationLabels(),
        (async () => {
          const fp = await loadFieldPermissionMap(sp);
          setFieldPermMap(fp);
          
          const roles = await loadGeneralRoleUsers(sp);
          setRoleUsers(prev => ({
            ...prev,
            FinancialAdvisor: roles.FinancialAdvisor.map(e => e.toLowerCase()),
            Lawyer: roles.Lawyer.map(e => e.toLowerCase()),
            PMOIntegrationTeam: roles.PMOIntegrationTeam.map(e => e.toLowerCase()),
          }));
        })(),
      ]);
      
      sessionStorage.removeItem(AUTO_RELOAD_KEY);
      await resetReloadGuard(sp);
      
    } catch (e: any) {
      // Only attempt reload guard if not in edit mode
      try {
        if (!isEditMode) {
          // ✅ Ensure SP is ready before reload guard operations
          await sp.web.select("Title")();
          const didReload = await markReloadedAndReload(sp);
          if (didReload) return;
        }
      } catch (guardErr) {
        console.log("Reload guard failed:", guardErr);
      }
      
      setMsg({ 
        type: MessageBarType.error, 
        text: '😋Error loading initial data: ' + (e?.message || e) + "\n Please reload the page and everything will be ok."
      });
    } finally {
      setBusy(false);
    }
  })();
}, [sp]);



  useEffect(() => {
    (async () => {
      if (!integrationId) return;
      try {
        console.log("🥯1");
        setBusy(true);
        console.log("🥯2");
        await loadFormForIntegration();
        console.log("🥯3");
      } catch (e: any) {
       console.log("🥯4");
        setMsg({ type: MessageBarType.error, text: '🥯Error loading the form: ' + (e?.message || e) });
         const msgText = String(e?.message || e || '');
        setMsg({
          type: MessageBarType.error,
          text: '🛼Error loading initial data: ' + msgText
        });
        console.log("1");

        // 🔁 אם זו בדיוק השגיאה של web, ננסה ריענון *פעם אחת בלבד*
        if (msgText.indexOf("Cannot read properties of undefined (reading 'web')")!= -1) {
          const alreadyTried = sessionStorage.getItem(AUTO_RELOAD_KEY);
          console.log("1");
          if (!alreadyTried) {
            sessionStorage.setItem(AUTO_RELOAD_KEY, '1');
            console.log("1");
            window.location.reload();
            console.log("1");
          }
        }
        console.log("1");
      } finally {
        setBusy(false);
      }
      console.log("1");
    })();
  }, [integrationId]);

  
  const onChangeField = (internal: string, value: any) => {
    console.log("⛔ allowed");
    if (internal === 'DecisionAppliesToOtherWorksTende') {
    const allowed = [
      'All Infra 1 tenders',
      'Not relevant to additional tenders',
      'Infra#1 DB - M3-WP2', 
      'Infra#1 DB - M2-WP3',
      'Infra#1 DB - M1-WP1 + WP2', 
      'M3-WPO (Outer Boxes)',
      'Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3', 
      'Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2',
      'Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes)', 
      'Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2',
      'Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2',
      'Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2', 
      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3',
      'Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes)',
      'M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2',
      'M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3',
      'M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2', 
      'Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2',
      'Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3',
      'Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3',
      'Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2',

      'Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2',
      'Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2',
      'Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2',
      'Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2',

      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3',
      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2',
      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2',
      'Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3',

      'M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3',
      'M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2',
      'M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2',
      'M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2',
      'M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2',
      'M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3',
      'Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2',
      'Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3',
      'Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2',
      'Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3',

      'Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2',
      'Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2',
      'Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2',
      'Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2',

      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3',
      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2',
      'Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3',
      'Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2',

      'M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2',
      'M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3',
      'M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2',
      'M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2',
      'M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3',
      'M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2',
      ''
    ];

    const vStr = String(value ?? '').trim();
    console.log("vStr 🦒 ", vStr);
    if (allowed.indexOf(vStr) === -1) {
      console.log("OI VEU 🦒😭🔮");
      console.warn('Blocked invalid value for DecisionAppliesToOtherWorksTende:', value);

      
      return; // ⛔ לא מעדכנים ל-draft
    }

   
  }

  // --- כאן האוטומציה החדשה ---
  setPmoDraft((prev: any) => {
    const next = { ...(prev || {}), [internal]: value };
    const nowIso = new Date().toISOString();

    // 1) ActualDate מתעדכן כש-StatusOfRFCresponseOrTcRFC משתנה
    if (internal === 'StatusOfRFCresponseOrTcRFC') {
      const prevVal = String(prev?.StatusOfRFCresponseOrTcRFC ?? '');
      const newVal  = String(value ?? '');
      if (prevVal !== newVal) {
        next.ActualDate = nowIso;
      }
    }

    // 2) DecisionDate מתעדכן כש-DecisionRegardingProposedChange משתנה
    if (internal === 'DecisionRegardingProposedChange') {
      const prevVal = String(prev?.DecisionRegardingProposedChange ?? '');
      const newVal  = String(value ?? '');
      if (prevVal !== newVal) {
        next.DecisionDate = nowIso;
      }
    }

    return next;
  });

    setPmoDraft((prev: any) => ({ ...prev, [internal]: value }));

      
    // אם השדה הפך ללא ריק – ננקה שגיאת "שדה חובה"
    setValidationErrors((prev) => {
      if (!prev[internal]) return prev;
      const copy = { ...prev };
      delete copy[internal];
      return copy;
    });
  };

  function isValidUrl(value: string): boolean {
  try {
    new URL(value);
    return true;
  } catch {
    return false;
  }
}

  const onSave = async (options?: { updateEditingDate?: boolean }) => {
    

    console.log("🐴🐴🐴 in onsave ");
    
    if (!pmoItem || !pmoItem.Id) return;
    // ---- ולידציית שדות חובה לפני השמירה ----
    const newErrors: Record<string, string> = {};
    let flage = false;
    for (let i = 0; i < REQUIRED_FIELDS.length; i++) {
      const internal = REQUIRED_FIELDS[i];

      // אם אין למשתמש הרשאה לערוך את השדה – הוא לא שדה חובה עבורו
      if (!canEditField(internal)) {
        continue;
      }

      if (!isFieldVisibleNow(internal)) {
        console.timeLog("🐴 !isFieldVisibleNow(internal) ", internal);
        continue;
      }

      
      console
      const v = pmoDraft ? pmoDraft[internal] : undefined;

      if(internal === "RFCresponseAsPublishedToBeFilled"){
        console.log("RFCresponseAsPublishedToBeFilled v ", v);
        if(v === false){
          flage = true;
        }
      }

      
      

      if(flage === true){
         if(internal === "Addendum"){
          continue;
        }
        if(internal === "RFCResponseLetterNo"){
          continue;
        }

        
        if(internal === "TenderCommitteeApprovalDate"){
          console.log("🐴🐴 internal === TenderCommitteeApprovalDate and flage === true");
          continue;
        }
      }

      
     

      
      let isEmpty = false;
      if (v === null || v === undefined) {
        console.log("🐴 internal ", internal);
        isEmpty = true;
      } else if (typeof v === 'string') {
        isEmpty = v.trim() === '';
      } else if (Array.isArray(v)) {
        isEmpty = v.length === 0;
      }

      if (isEmpty) {
        newErrors[internal] = 'Required field';
      }else if(!isEmpty){
        const info = pmoFieldInfoMap[internal];
        const t = String(info?.TypeAsString || '').toLowerCase();

        if (t === 'url' || t === 'hyperlink') {
        let urlToCheck = "";

        if (typeof v === "string") {
          urlToCheck = v.trim();
        }
        else if (typeof v === "object" && v !== null) {
          urlToCheck = String(v.Url || v.url || "").trim();
        }

        if (urlToCheck && !isValidUrl(urlToCheck)) {
          newErrors[internal] = "Invalid URL format";
        }
      }

      }
    }

    if (Object.keys(newErrors).length > 0) {
      setValidationErrors(newErrors);
      console.log("🐴🐴🐴🐴 msg ", msg);
      setMsg({
        type: MessageBarType.error,
        text: 'Please fill in all required fields marked in red before saving.'
      });
      return; // לא ממשיכים לשמירה ב־SharePoint
    }


        try {
        setBusy(true);

        // נבנה עותק שנשמור בפועל
        const draftToSave: any = { ...(pmoDraft || {}) };
        if (options?.updateEditingDate) {
          draftToSave.IntegrationTeamDecisionEditingLa =
            new Date().toISOString();
        }

        // המרה של שדה ההיפר־לינק לפורמט של SharePoint:
        // { Url: 'https://...', Description: '...' }
       

        // 💡 Normalize all URL / Hyperlink fields before saving
        for (const internal in pmoFieldInfoMap) {
          if (!Object.prototype.hasOwnProperty.call(pmoFieldInfoMap, internal)) continue;

          const info = pmoFieldInfoMap[internal];
          const t = String(info?.TypeAsString || '').toLowerCase();

          // גם URL וגם Hyperlink
          if (t !== 'url' && t !== 'hyperlink') continue;


          const val = draftToSave[internal];
          if (val === undefined || val === null) {
            draftToSave[internal] = null;
            continue;
          }

          // אם כבר בפורמט הנכון – לא נוגעים
          if (typeof val === 'object' && (val.Url || val.url)) continue;

          const urlStr = String(val || '').trim();
          if (!urlStr) {
            draftToSave[internal] = null;
          } else {
            draftToSave[internal] = {
              Url: urlStr,
              Description: urlStr,
            };

          }
        }
        

        
        console.log('🧾 draftToSave before save:', draftToSave);

        const saved = await savePmoItem(sp, PMO_LIST_ID, pmoItem.Id, draftToSave);
       

        setPmoItem(saved);
        setPmoDraft(saved);
        setMsg({ type: MessageBarType.success, text: 'Saved successfully.' });
        setValidationErrors({});

      // ---- אם אין שגיאות חובה, ממשיכים לשמור ----
      

        // 🆕 אחרי שה-PMO נשמר בהצלחה – נסנכרן גם ל-INTEGRATION
        try {
          if (integrationId) {
            //await syncPmoToIntegration(sp, integrationId, pmoDraft);
            await syncPmoToIntegration(sp, integrationId, draftToSave);
            console.log('✅ Synced PMO → INTEGRATION for item', integrationId);
          } else {
            console.warn('syncPmoToIntegration: integrationId is null – no sync done');
          }
        } catch (syncErr) {
          console.error('❌ Failed to sync PMO → INTEGRATION', syncErr);
          // לא מפילים למשתמש את השמירה – זה רק סנכרון עזר
        }
        console.log("🏔️🏔️🏔️ integrationItem ", integrationItem);
         console.log("🏔️🏔️🏔️ draftToSave ", draftToSave);
       await splitTenderAndCreateIntegrationItems({
        sp,
        integrationListId: '2c962132-409d-4bf2-9440-3b3b6c7975a0',
        pmoItem: draftToSave, // או draftToSave אם את בטוחה שהוא מלא ונקי
        itegrationItem: integrationItem,
        tenderSourceInternalName: "DecisionAppliesToOtherWorksTende", // או "TenderNumber"
        pmoToIntegrationMap: PMO_TO_INTEGRATION_FIELD_MAP,
        linkFieldInternalName: "NTA_x2019_s_x0020_reference", // אם קיים אצלכם
        linkValue: draftToSave?.IntegrationId, // או saved?.Id — לפי מה שהשדה מצפה
      });



      } catch (e: any) {
        setMsg({ type: MessageBarType.error, text: 'Save failed: ' + (e?.message || e) });
      } finally {
        setBusy(false);
      }

      

  };

 

  function formatDateDDMMYYYY(value: any): string {
  if (!value) return '';
  const d = new Date(value);
  if (isNaN(d.getTime())) return String(value);
  const dd = ('0' + d.getDate()).slice(0 - 2);
  const mm = ('0' + (d.getMonth() + 1)).slice(0 - 2);
  const yyyy = d.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
}


  const renderIntegrationReadonly = () => {

    if (!integrationItem) return null;

        // 🧩 אם יש View מוגדר – ניקח את הסדר ממנו; אחרת – כל השדות מהפריט
    const allKeys = Object.keys(integrationItem);

    // אם יש View – נשתמש בשדות מה-View; אחרת – כל השדות מהפריט
    const baseKeys = (integrationViewFieldOrder && integrationViewFieldOrder.length
      ? integrationViewFieldOrder //.filter(k => allKeys.indexOf(k)!= -1)
      : allKeys
    );

    console.log("integrationViewFieldOrder ", integrationViewFieldOrder);

  

    const keys = baseKeys.filter(k => {
      console.log(" - integration key ", k)
      if (isSystemField(k)){ 
        console.log("🛹 is system");
        return false;}

      if (k === 'Id' || k === 'ID' || k === 'Title') return false;
      const info = integrationFieldInfoMap[k];
      
      if (info && info.Hidden === true) {
        console.log("🙈k info.Hidden === true ");
        return false;}
      return true;
    });
 
    return (
      <Stack tokens={{ childrenGap: 10 }}>
        {keys.map(k => {
  const raw = integrationItem[k];
  const info = integrationFieldInfoMap[k];

  let text: string;

const typeStr = String(info?.TypeAsString || '').toLowerCase();
  if(k === "RelatedBidderRFCsexist_x003f_" || k === "weretheRelatedBidderRFCsresponde"){
    text = '';
  }

  // 🎯 1. Boolean → Yes / No
  if (typeStr === 'boolean') {

    if (raw === true) {
      text = 'Yes';
    } else if (raw === false) {
      text = 'No';
    } else {
      text = ''; // במידה וריק/לא מוגדר
    }
    console.log("🎯 k ", k);
  }
  // 🎯 1. אם זה שדה תאריך (כולל ApplicationDate) – פורמט יפה
  else if (info && String(info.TypeAsString || '').toLowerCase().indexOf('date') > -1 && raw) {
      text = formatDateDDMMYYYY(raw);
    }
    // 🎯 2. אם זה HTML – ננקה לתצוגת טקסט
    else if (typeof raw === 'string' && /<[^>]+>/.test(raw)) {
      text = htmlToPlainText(raw);
    }
    // 🎯 3. כל השאר – סתם להמיר למחרוזת
    else {
      text = String(raw ?? '');
    }


    return (
      <div
        key={k}
        style={{
          padding: 10,
          background: 'rgba(15,23,42,0.02)',
          borderRadius: 12,
          border: '1px solid #e5e7eb',
          boxShadow: '0 6px 18px rgba(15,23,42,0.06)'
        }}
      >
        <div style={{ fontSize: 12, color: '#6b7280', marginBottom: 4 }}>
          {integrationLabels[k] || k}
        </div>
        <div style={{ fontWeight: 600, whiteSpace: 'pre-wrap', color: '#111827' }}>{text}</div>
      </div>
    );
  })}

      </Stack>
    );
  };

  // --- לוגיקת תלות לשדה IntegrationTeamDecisionImplement ---
  const APPLIES_FIELD = 'DecisionAppliesToOtherWorksTende';
  const TARGET_FIELD = 'IntegrationTeamDecisionImplement';
  const DATE_FIELD = 'DateForIntegrationTeamDecisionIm';

  // האם "Applies" מאפשר עריכה?
  const appliesVal = String(pmoDraft?.[APPLIES_FIELD] || '').trim();
  const canEdit_ITDI =
   (appliesVal === 'All Infra 1 tenders' || appliesVal === 'Infra#1 DB - M3-WP2'|| appliesVal === 'Infra#1 DB - M2-WP3'
  || appliesVal === 'Infra#1 DB - M1-WP1 + WP2' || appliesVal ==='M3-WPO (Outer Boxes)' ||
  appliesVal.indexOf('All Infra 1 tenders') !== -1 ||
   appliesVal.indexOf('Infra#1 DB - M3-WP2') !== -1 || appliesVal.indexOf('Infra#1 DB - M2-WP3') !== -1 ||
    appliesVal.indexOf('Infra#1 DB - M1-WP1 + WP2') !== -1 || appliesVal.indexOf('M3-WPO (Outer Boxes)') !== -1);



  console.log("🦧 canEdit_ITDI ", canEdit_ITDI, " appliesVal ", appliesVal); 
  // הפקת תאריך לתצוגה
  let dateText = '';
  try {
    const rawDate = pmoDraft?.[DATE_FIELD];
    if (rawDate) {
      const d = new Date(rawDate);
      if (!isNaN(d.getTime())) {
        dateText = d.toLocaleDateString('he-IL');
        console.log("🥨dateText ", dateText);
      }
    }
  } catch { }
   // לפני ה־itdiOptions, פעם אחת:
    const tenderPStr = String(integrationItem?.TenderPhase || '').trim().toLowerCase();
    const isPhase2 =
      tenderPStr.indexOf('phase 2') !== -1; 
    console.log("🥗canEdit_ITDI ", canEdit_ITDI, " pmoDraft?.['StatusOfRFCresponseOrTcRFC'] === 'Issued' ",  pmoDraft?.['StatusOfRFCresponseOrTcRFC'] === 'Issued', " isPhase2 ", isPhase2, " " );
    console.log("🥨dateText ", dateText);
    const itdiOptions = (canEdit_ITDI && ((isPhase2 &&  pmoDraft?.['StatusOfRFCresponseOrTcRFC'] === 'Issued') ||(!isPhase2)))
    ? [
      { key: 'Done', text: dateText ? `Done — ${dateText}` : 'Done' },
      { key: 'Pending', text: dateText ? `Pending — ${dateText}` : 'Pending' }
    ]:canEdit_ITDI ? [
      { key: 'Pending', text: dateText ? `Pending — ${dateText}` : 'Pending' }
    ]
    : [
      { key: 'Not required', text: 'Not required' },
    ];
  
  useEffect(() => {
    const v = String(pmoDraft?.DecisionAppliesToOtherWorksTende ?? '').trim();
    if (!v) return;
    //const vDecisionRegardingProposedChange = String(pmoDraft?.DecisionRegardingProposedChange??'').trim();
    let allowed = [
      'All Infra 1 tenders',
      'Not relevant to additional tenders',
      'Infra#1 DB - M3-WP2', 
      'Infra#1 DB - M2-WP3',
      'Infra#1 DB - M1-WP1 + WP2', 
      'M3-WPO (Outer Boxes)',
      'Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3', 
      'Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2',
      'Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes)', 
      'Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2',
      'Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2',
      'Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2', 
      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3',
      'Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes)',
      'M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2',
      'M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3',
      'M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2', 
      'Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2',
      'Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3',
      'Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3',
      'Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2',

      'Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2',
      'Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2',
      'Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2',
      'Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2',

      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3',
      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2',
      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2',
      'Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3',

      'M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3',
      'M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2',
      'M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2',
      'M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2',
      'M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2',
      'M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3',
      'Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2',
      'Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3',
      'Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2',
      'Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3',

      'Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2',
      'Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2',
      'Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2',
      'Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2',

      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3',
      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2,M3-WPO (Outer Boxes)',
      'Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3,M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2',
      'Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3',
      'Infra#1 DB - M1-WP1 + WP2,M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2',

      'M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2',
      'M3-WPO (Outer Boxes),Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3',
      'M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2,Infra#1 DB - M1-WP1 + WP2',
      'M3-WPO (Outer Boxes),Infra#1 DB - M2-WP3,Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2',
      'M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M3-WP2,Infra#1 DB - M2-WP3',
      'M3-WPO (Outer Boxes),Infra#1 DB - M1-WP1 + WP2,Infra#1 DB - M2-WP3,Infra#1 DB - M3-WP2',
      ''
    ];
   
    

    if (allowed.indexOf(v) === -1) {
      console.warn('Auto-fixing invalid DecisionAppliesToOtherWorksTende:', v);

      setPmoDraft((prev:any) => ({
        ...prev,
        DecisionAppliesToOtherWorksTende: null
      }));

      setMsg({
        type: MessageBarType.warning,
        text: `Field "Decision Applies" had invalid value "${v}" and was cleared.`
      });
    }
  }, [pmoDraft?.DecisionAppliesToOtherWorksTende]);

  useEffect(() => {
    console.log("🍢 canEdit_ITDI ", canEdit_ITDI);
    if (!canEdit_ITDI) {
      console.log("🍥🎀🥩 canEdit_ITDI is not null - TARGET_FIELD", TARGET_FIELD, " pmoDraft?.[TARGET_FIELD] ", pmoDraft?.[TARGET_FIELD]);
      if (pmoDraft?.[TARGET_FIELD] !== 'Not required') {
        console.log("🪂");
        setPmoDraft((prev: any) => ({ ...(prev || {}), [TARGET_FIELD]: 'Not required' }));
      }
    }
    console.log("in the end 🍢 canEdit_ITDI ", canEdit_ITDI);
  }, [canEdit_ITDI, pmoDraft?.[TARGET_FIELD]]);


  const dynamicHideFieldstenders = React.useMemo<string[]>(() => {
    const base = ['Integration', 'IntegrationId', 'Id', 'ID', 'Title'];

    // 1. יישור לפי ה-View: כל שדה שלא מופיע ב-View → מוסתר
    if (viewFieldOrder && viewFieldOrder.length && pmoDraft) {
      const allFields = Object.keys(pmoDraft);
      for (let i = 0; i < allFields.length; i++) {
        const f = allFields[i];
        /*
        if (f === 'RFCorTcRFCasPublishedByNTaNew') continue;*/
        if (base.indexOf(f) === -1 && viewFieldOrder.indexOf(f) === -1) {
          base.push(f);
        }
      }
    }


  // 2. לוגיקה של RevisionIncludesChangeInTenderDo
  const rv = pmoDraft?.RevisionIncludesChangeInTenderDo;
  const isYes = (() => {
    if (typeof rv === 'boolean') return rv;
    const s = String(rv ?? '').trim().toLowerCase();
    return s === 'y' || s === 'yes' || s === 'true' || s === '1';
  })();

  if (!isYes) base.push('RFCResponseLetterNo');

  // 3. 🔹 לוגיקה חדשה לפי sentProtocol
  const rawSent = pmoDraft?.sentProtocol;
  const sentStr = String(rawSent ?? '').trim();

  // אם זה עמודת בחירה/טקסט עם "כן"
  const isSentYes =
    rawSent === true ||          // אם זה Yes/No (boolean)
    sentStr === 'כן' ||          // אם זה טקסט בעברית
    sentStr.toLowerCase() === 'yes'; // אם יבוא לך באנגלית בעתיד

  // אם זה *לא* "כן" → מסתירים Addendum ו-IntegrationTeamDecisionImplement
  if (!isSentYes) {
    if (base.indexOf('Addendum') === -1) base.push('Addendum');
    //if (base.indexOf(TARGET_FIELD) === -1) base.push(TARGET_FIELD);
  }

  const tenderPhaseRaw = pmoDraft?.TenderPhase;
  const tenderPhaseStr = String(tenderPhaseRaw ?? '').toLowerCase();

  // מספיק שאו שזה בדיוק "1" או שמופיע בו "phase 1"
  const isPhase1 = tenderPhaseStr === 'phase 1 - bidders’ requests for clarifications (rfcs) of tender documents' && tenderPhaseStr.indexOf('phase 1') != -1;

  if (isPhase1 && !base.includes('StatusOfRFCresponseOrTcRFC')) {
    base.push('StatusOfRFCresponseOrTcRFC');
  }



  return base;
}, [pmoDraft, pmoDraft?.RevisionIncludesChangeInTenderDo, pmoDraft?.sentProtocol, viewFieldOrder]);


// ✅ אילו שדות של Tender Team *כן* יוצגו (TENDER_TEAM_FIELDS פחות מה שמוסתר דינאמית)
const tenderTeamVisibleFields: string[] = React.useMemo(() => {
  const dynHidden = dynamicHideFieldstenders || [];
  return TENDER_TEAM_FIELDS.filter(f => !dynHidden.includes(f));
}, [dynamicHideFieldstenders]);

// ✅ hideFields עבור ה־EditableFields של Tender Team:
// מסתיר *כל* השדות שלא נמצאות ב־tenderTeamVisibleFields
const tenderTeamHideFields: string[] = React.useMemo(() => {
  const src = (pmoDraft || pmoItem) ?? {};
  return Object.keys(src).filter(k => !tenderTeamVisibleFields.includes(k));
}, [pmoDraft, pmoItem, tenderTeamVisibleFields]);


const dynamicHideFields = React.useMemo<string[]>(() => {
  const base = ['Integration', 'IntegrationId', 'Id', 'ID', 'Title'];

  // 1. יישור לפי ה-View: כל שדה שלא מופיע ב-View → מוסתר
  if (viewFieldOrder && viewFieldOrder.length && pmoDraft) {
    const allFields = Object.keys(pmoDraft);
    for (let i = 0; i < allFields.length; i++) {
      const f = allFields[i];
      /*
      if (f === 'RFCorTcRFCasPublishedByNTaNew') continue;*/
      if (base.indexOf(f) === -1 && viewFieldOrder.indexOf(f) === -1) {
        base.push(f);
      }
    }
  }

  for(let i = 0; i < TENDER_TEAM_FIELDS.length; i++){
    const f = TENDER_TEAM_FIELDS[i];
    base.push(f);
  }

  // 2. לוגיקה של RevisionIncludesChangeInTenderDo
  const rv = pmoDraft?.RevisionIncludesChangeInTenderDo;
  const isYes = (() => {
    if (typeof rv === 'boolean') return rv;
    const s = String(rv ?? '').trim().toLowerCase();
    return s === 'y' || s === 'yes' || s === 'true' || s === '1';
  })();

  if (!isYes) base.push('RFCResponseLetterNo');

  // 3. 🔹 לוגיקה חדשה לפי sentProtocol
  const rawSent = pmoDraft?.sentProtocol;
  const sentStr = String(rawSent ?? '').trim();

  // אם זה עמודת בחירה/טקסט עם "כן"
  const isSentYes =
    rawSent === true ||          // אם זה Yes/No (boolean)
    sentStr === 'כן' ||          // אם זה טקסט בעברית
    sentStr.toLowerCase() === 'yes'; // אם יבוא לך באנגלית בעתיד

  // אם זה *לא* "כן" → מסתירים Addendum ו-IntegrationTeamDecisionImplement
  if (!isSentYes) {
    if (base.indexOf('Addendum') === -1) base.push('Addendum');
    if (base.indexOf(TARGET_FIELD) === -1) base.push(TARGET_FIELD);
  }

  return base;
}, [pmoDraft, pmoDraft?.RevisionIncludesChangeInTenderDo, pmoDraft?.sentProtocol, viewFieldOrder]);



  const isFieldVisibleNow = (internal: string): boolean => {
    if (['Integration', 'IntegrationId', 'Id', 'ID', 'Title'].indexOf(internal) > -1) {
      return false;
    }

    if (dynamicHideFields.indexOf(internal) > -1) {
      return false;
    }
      const sentVal = String(pmoDraft?.sentProtocol ?? '').trim();
      const isSentYes = (sentVal === 'true');
      console.log("💛sentVal ", sentVal);

      // אם זה אחד השדות הרלוונטיים, והוא לא "כן" → להסתיר
      if ((internal === 'Addendum' || internal === TARGET_FIELD) && !isSentYes) {
        return false;
      }

    const decision = String(pmoDraft?.DecisionRegardingProposedChange || '').trim();
    const revInc = String(pmoDraft?.RevisionIncludesChangeInTenderDo || '').trim().toLowerCase();
    const tenderPhaseStr = String(integrationItem?.TenderPhase || '').trim().toLowerCase();
        // האם RevisionIncludesChangeInTenderDo הוא "כן"/True
    const revIsTrue = (() => {
      const rv = pmoDraft?.RevisionIncludesChangeInTenderDo;
      if (typeof rv === 'boolean') return rv;
      const s = String(rv ?? '').trim().toLowerCase();
      return s === 'true' || s === 'y' || s === 'yes' || s === '1';
    })();



    if (internal === 'RFCresponseAsPublishedToBeFilled' || internal === 'StatusOfRFCresponseOrTcRFC') {
      if (decision !== 'Accept') return false;
    }

    if (internal === 'Addendum' || internal === 'TenderCommitteeApprovalDate') {
      const isNo = (revInc === 'n' || revInc === 'no' || revInc === 'false' || revInc === '0');
      if (isNo) {
        console.log("🐴🐴🐴🐴🐴 internal === 'TenderCommitteeApprovalDate' and isNo false");
        return false;
      }

      if (!revIsTrue || !isPhase2) {
        console.log("🐴🐴🐴🐴  isFieldVisibleNow TenderCommitteeApprovalDate");
        return false;
      }
    }

    if (internal === 'dog') {
      return false;
    }

    if (internal === 'RevisedWordingFinalForPublicatio') {
      if (tenderPhaseStr !== 'phase 1 – preparation of tender documents') return false;
      //return false;
    }

    if (
      internal === 'RevisionIncludesChangeInTenderDo' ||
      internal === 'RFCResponseLetterNo' ||
      internal === 'RFCresponseAsPublishedToBeFilled'
    ) {
      if (tenderPhaseStr !== 'phase 2 - bidders’ requests for clarifications (rfcs) of tender documents') {
        return false;
      }
    }
    if(internal === "formCreator"){
      return false;
    }

    return true;
  };

  const isFieldInTender = (internal: string): boolean => {
    if (TENDER_TEAM_FIELDS.indexOf(internal) > -1) {
      return false;
    }
    return true;
  };



  const canEditField = (internal: string): boolean => {
    console.log("🧨internal ", internal);
    const emailLc = String(me?.Email || '').toLowerCase();
    if (!emailLc) return false;

    const olm = getOriginatingLineManager();
    console.log("🧨internal ", internal);
    if (internal === TARGET_FIELD) {
      console.log("🧨🎇 canEdit_ITDI ", canEdit_ITDI);
      return canEdit_ITDI && canUserEditField(emailLc, internal, fieldPermMap, {
        IntegrationTeam: roleUsers.IntegrationTeam,
        IntegrationTeamLawyer: roleUsers.IntegrationTeamLawyer,
        FinancialAdvisorIntegrationTeam: roleUsers.FinancialAdvisorIntegrationTeam,
        M1TenderTeam: roleUsers.M1TenderTeam,
        LawyerM1TenderTeam: roleUsers.LawyerM1TenderTeam,
        M2TenderTeam: roleUsers.M2TenderTeam,
        LawyerM2TenderTeam: roleUsers.LawyerM2TenderTeam,
        M3TenderTeam: roleUsers.M3TenderTeam,
        LawyerM3TenderTeam: roleUsers.LawyerM3TenderTeam,
        FinancialAdvisor: roleUsers.FinancialAdvisor,
        Lawyer: roleUsers.Lawyer,
        PMOIntegrationTeam: roleUsers.PMOIntegrationTeam,
        PMOTenderTeam: roleUsers.PMOTenderTeam
      },
      olm
    );
    }

    return canUserEditField(emailLc, internal, fieldPermMap, roleUsers, olm);
  };



  const renderPmoEditableBySteps = () => {
      // 🐛 DEBUG – print all PMO fields & visibilityte
  if (pmoDraft) {
    const allFields = Object.keys(pmoDraft || {});

    console.log('PMO all fields (internal names):', allFields);
    console.log('PMO viewFieldOrder:', viewFieldOrder);
    console.log('PMO dynamicHideFields:', dynamicHideFields);

    const temp = allFields.filter(f => isFieldVisibleNow(f));
    console.log("🦁temp ", temp);
    const visibleNow = temp.filter(f=> isFieldInTender(f));
    console.log("🔮 visibleNow ",visibleNow );
    console.log('PMO fields that isFieldVisibleNow() == true:', visibleNow);

    
  }

  // 🔹 נחלץ את המפתח של ה־Phase מתוך TenderPhase
  const phaseKey = getPhaseViewKeyFromTenderPhase(integrationItem?.TenderPhase);

  const allSteps = Object.keys(stepsInternal);

  // 🔹 אם יש Phase מזוהה – נסנן רק את ה־steps שהשם שלהם מתחיל בו
  const stepNames = phaseKey
    ? allSteps.filter(s => s.toLowerCase().indexOf(phaseKey)===0)
    : allSteps;

  // אם אין בכלל steps – מציגים טופס מלא בלי Pivot
  if (!stepNames.length) {
    const order = viewFieldOrder && viewFieldOrder.length
        ? viewFieldOrder
        : Object.keys(pmoDraft || {});
    return (
      <EditableFields
        item={pmoDraft}
        onChange={onChangeField}
        fieldOrder={order}//stepsInternal[step] || []}
        hideFields={dynamicHideFields}
        internalToTitle={pmoLabels}
        fieldInfoMap={pmoFieldInfoMap}
        canEdit={canEditField}
        placeholderMap={placeholders}
        choiceOverrides={{
          [TARGET_FIELD]: itdiOptions
        }}
        tenderPhase={String(integrationItem?.TenderPhase || '')}
        requiredMap={requiredMap}
        errorMap={validationErrors}
        labelOverrides={PMO_LABEL_OVERRIDES}
      />
    );
  }
  const getStepFieldOrder = (step: string): string[] => {
  const stepFields = stepsInternal[step] || [];

  // אם יש View – נכבד אותו קודם, ואז נסנן לפי ה-step
  let order: string[];
  if (viewFieldOrder && viewFieldOrder.length) {
    if (stepFields.length) {
      order = viewFieldOrder.filter(f => stepFields.indexOf(f) !== -1);
    } else {
      order = viewFieldOrder;
    }
  } else {
    // בלי View – נשאר רק עם ה-step המקורי
    order = stepFields;
  }

  return order;//mapped;
};


  // לוודא שה־selectedKey תמיד שייך ל־stepNames
  const effectiveActiveStep = stepNames.indexOf(activeStep)!= -1
    ? activeStep
    : (stepNames[0] || '');

  return (
    <Pivot
      selectedKey={effectiveActiveStep}
      onLinkClick={(i) => setActiveStep(i?.props.itemKey || '')}
      styles={{
        root: {
          marginTop: 8
        },
        link: {
          fontWeight: 600,
          fontSize: 14
        }
      }}
    >
      {stepNames.map(step => (
        <PivotItem headerText={step} itemKey={step} key={step}>
          <div style={{ marginTop: 10 }}>
            <EditableFields
              item={pmoDraft}
              onChange={onChangeField}
              fieldOrder={getStepFieldOrder(step)}
              hideFields={dynamicHideFields}
              internalToTitle={pmoLabels}
              fieldInfoMap={pmoFieldInfoMap}
              canEdit={canEditField}
              placeholderMap={placeholders}
              choiceOverrides={{
                [TARGET_FIELD]: itdiOptions
              }}
              tenderPhase={String(integrationItem?.TenderPhase || '')}
              requiredMap={requiredMap}
              errorMap={validationErrors}
              labelOverrides={PMO_LABEL_OVERRIDES}
            />
          </div>
        </PivotItem>
      ))}
    </Pivot>
  );
};


console.log("🔎 integrationSearch ", integrationSearch, " integrationId ", integrationId);

  return (
    <div
      dir="ltr"
      style={{
        background: PAGE_BG,
        minHeight: '100vh',
        padding: '24px 32px',
        textAlign: 'left'
      }}
    >
      <Stack
        tokens={{ childrenGap: 20 }}
        styles={{
          root: {
            maxWidth: 1400,
            margin: '0 auto'
          }
        }}
      >
        {}
        <Stack
          horizontal
          horizontalAlign="space-between"
          verticalAlign="center"
          tokens={{ childrenGap: 12 }}
        >
          <Stack tokens={{ childrenGap: 4 }}>
            <span
              style={{
                fontSize: 13,
                color: '#6b7280',
                background: ACCENT_SOFT,
                padding: '4px 12px',
                borderRadius: 999,
                alignSelf: 'flex-start'
              }}
            >
              PMO • Smart Decision Form
            </span>
            <h1 style={{ margin: 0, fontSize: 28, color: '#0f172a' }}>
              Application Form to Integration Team – LM's proposed changes /  deviations from precedent 
            </h1>
            <p style={{ margin: 0, fontSize: 14, color: '#6b7280', maxWidth: 620 }}>
              Select an item on the left, and get a role-specific PMO form on the right, with highlighted required fields and a clear workflow.
            </p>
          </Stack>

          {}
          
        </Stack>

        {msg ? (
          <MessageBar
            messageBarType={msg.type}
            styles={{ root: { borderRadius: 10, boxShadow: '0 10px 30px rgba(15,23,42,0.08)' } }}
          >
            {msg.text}
          </MessageBar>
        ) : null}

        {}
        <div
          style={{
            background: CARD_BG,
            borderRadius: CARD_RADIUS,
            boxShadow: CARD_SHADOW,
            padding: 18,
            border: '1px solid rgba(148,163,184,0.25)'
          }}
        >
          <Stack horizontal tokens={{ childrenGap: 12 }} style={{ alignItems: 'flex-end' }}>
            <Stack grow tokens={{ childrenGap: 8 }}>
              <span
                style={{
                  fontSize: 13,
                  color: '#64748b'
                }}
              >
                _
              </span>
              <div style={{ fontSize: 18, fontWeight: 600, color: '#0f172a' }}>
                Tender item selection
              </div>

              {}
              <div>
                My OLM: {myRoles.join(", ") || "—"}
              </div>
              <Dropdown
                label="Filter by Originating Line Manager"
                selectedKey={olmFilter}
                onChange={(_, opt) => setOlmFilter((opt?.key as string) || 'ALL')}
                options={[
                  { key: 'ALL', text: 'All line managers' },
                  { key: 'M1', text: 'M1' },
                  { key: 'M2', text: 'M2' },
                  { key: 'M3', text: 'M3' },
                ]}
                styles={{ root: { maxWidth: 220 } }}
              />
            
              <ComboBox
                label="Search Integration by NTA reference"
                placeholder={busy ? 'Loading…' : 'Start typing to search…'}
                options={filteredIntegrationChoices}
                selectedKey={integrationId ?? undefined}
                text={integrationSearch}
                autoComplete="on"
                allowFreeform
                openOnKeyboardFocus
                useComboBoxAsMenuWidth
                onFocus={() => {
                  setIntegrationId(null);
                  setIntegrationSearch('');
                  //setComboOpen(true);

                }}
                onInputValueChange={(text?: string) => {
                  const t = text ?? '';
                  setIntegrationSearch(t);

                  if (!t.trim()) {
                    setIntegrationId(null);
                    return;
                  }

                  if (integrationId !== null) {
                    const chosen = integrationChoices.find(o => o.key === integrationId);
                    if (chosen && String(chosen.text || '') !== t) {
                      setIntegrationId(null);
                    }
                  }
                }}

                onClick={() => {
                  setIntegrationId(null);
                  setIntegrationSearch('');
                  //setComboOpen(true);
                }}

                onChange={(_, opt, __, value) => {
                  if (opt) {
                    setIntegrationId(opt.key as number);
                    setIntegrationSearch(opt.text ?? '');
                  } else {
                    setIntegrationId(null);
                    setIntegrationSearch(value ?? '');
                  }
                }}
                styles={{
                  root: { minWidth: 540 },
                  label: { fontWeight: 600 }
                }}
              />
            </Stack>

            <DefaultButton
              text="Refresh list"
              onClick={loadIntegrationChoices}
              disabled={busy}
              styles={{
                root: { borderRadius: 999, paddingInline: 18 },
              }}
            />
          </Stack>
        </div>

        {}
        <Stack
          horizontal
          wrap
          tokens={{ childrenGap: 24 }}
          styles={{ root: { alignItems: 'flex-start' } }}
        >
          {/* כרטיס שמאלי – Integration Display */}
          <Stack
            grow
            styles={{
              root: {
                background: CARD_BG,
                borderRadius: CARD_RADIUS,
                boxShadow: CARD_SHADOW,
                padding: 18,
                border: '1px solid rgba(148,163,184,0.25)',
                minWidth: 0
              }
            }}
            tokens={{ childrenGap: 10 }}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline' }}>
              <h3 style={{ margin: 0, fontSize: 18, color: ACCENT }}>Integration – Display</h3>
              <span style={{ fontSize: 12, color: '#9ca3af' }}>Read only</span>
            </div>
            <p style={{ margin: '0 0 8px', fontSize: 13, color: '#6b7280' }}>
            </p>
            {renderIntegrationReadonly()}
          </Stack>

          {/* כרטיס ימני – Integration Team Decision – Editing */}
          <Stack
            grow
            styles={{
              root: {
                background: CARD_BG,
                borderRadius: CARD_RADIUS,
                boxShadow: CARD_SHADOW,
                padding: 18,
                border: '1px solid rgba(148,163,184,0.25)',
                minWidth: 0
              }
            }}
            tokens={{ childrenGap: 10 }}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline' }}>
              <h3 style={{ margin: 0, fontSize: 18, color: ACCENT }}>Integration Team Decision – Editing</h3>
              <span style={{ fontSize: 12, color: '#9ca3af' }}>
                Required fields are marked in red • Permissions by role
              </span>
            </div>

            {renderPmoEditableBySteps()}
            <div style={{ marginTop: 16, display: 'flex', justifyContent: 'flex-end' }}>
              <PrimaryButton
                text={busy ? 'Saving…' : 'Save Integration Decision'}
                disabled={busy || !pmoItem}
                onClick={() => onSave({ updateEditingDate: false})}//true })}
                styles={{
                  root: {
                    borderRadius: 999,
                    paddingInline: 24,
                    fontWeight: 600,
                  },
                }}
              />
            </div>

          </Stack>
        </Stack>

        {/* 🔥 כרטיס נפרד – Tender Team status (מתחת לשני הכרטיסים למעלה) */}
        <Stack
          styles={{
            root: {
              background: CARD_BG,
              borderRadius: CARD_RADIUS,
              boxShadow: CARD_SHADOW,
              padding: 18,
              border: '1px solid rgba(148,163,184,0.25)',
              marginTop: 16
            }
          }}
          tokens={{ childrenGap: 10 }}
        >
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline' }}>
            <h3 style={{ margin: 0, fontSize: 18, color: ACCENT }}>
              Tender Team status
            </h3>
            <span style={{ fontSize: 12, color: '#9ca3af' }}>
              These fields are managed by the Tender Team and saved to the same PMO item.
            </span>
          </div>

          <EditableFields
            item={pmoDraft || pmoItem || {}}
            onChange={(internalName: string, value: any) => {
              setPmoDraft((prev: any) => ({
                ...(prev || pmoItem || {}),
                [internalName]: value,
              }));
            }}
            fieldInfoMap={pmoFieldInfoMap}
            fieldOrder={TENDER_TEAM_FIELDS}
            hideFields={tenderTeamHideFields}
            labelOverrides={PMO_LABEL_OVERRIDES}  
            canEdit={canEditField}
            tenderPhase={String(integrationItem?.TenderPhase || '')}
            choiceOverrides={{
              [TARGET_FIELD]: itdiOptions
            }}
            placeholderMap={placeholders}
          />
        </Stack>

        { }
        <Stack
          horizontal
          tokens={{ childrenGap: 10 }}
          horizontalAlign="space-between"
          verticalAlign="center"
        >
          <span style={{ fontSize: 12, color: '#6b7280' }}>
            Form changes are saved directly to SharePoint when you click<strong>Save</strong>.
          </span>
          <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
            <PrimaryButton
              text={busy ? 'Saveing' : 'Save'}
              onClick={() => onSave({ updateEditingDate: false })} 
              disabled={busy || !pmoItem}
              styles={{
                root: { borderRadius: 999, paddingInline: 26, fontWeight: 600 }
              }}
            />
            <DefaultButton
              text="Refresh Form"
              onClick={loadFormForIntegration}
              disabled={busy || !integrationId}
              styles={{
                root: { borderRadius: 999, paddingInline: 18 }
              }}
            />

            {msg && (
              <span
                style={{
                  marginLeft: 12,
                  padding: '4px 10px',
                  borderRadius: 999,
                  fontSize: 13,
                  fontWeight: 600,
                  backgroundColor:
                    msg.type === MessageBarType.success ? '#ecfdf3' : '#fef2f2',
                  color:
                    msg.type === MessageBarType.success ? '#166534' : '#b91c1c',
                  border: `1px solid ${
                    msg.type === MessageBarType.success ? '#bbf7d0' : '#fecaca'
                  }`,
                  whiteSpace: 'nowrap',
                }}
              >
                {msg.text}
              </span>
            )}
          </Stack>
        </Stack>
      </Stack>
    </div>
  );


};




export default FormApp;



