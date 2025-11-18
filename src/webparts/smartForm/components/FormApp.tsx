
// src/webparts/smartForm/components/FormApp.tsx
import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import {
  Stack, Pivot, PivotItem,
  PrimaryButton, DefaultButton,
  MessageBar, MessageBarType,
  Dropdown, IDropdownOption, TextField
} from '@fluentui/react';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { SPFI } from '@pnp/sp';
import { getSP } from './pnpjsConfig';
import { isSystemField } from '../shared/constants';
import { INTEGRATION_LIST_ID, INTEGRATION_PREVIEW_FIELD_INTERNAL } from '../shared/internalNames';
import {
  getFieldMapsByTitle,
  getFieldInfoMap,
  fetchIntegrationItemByGuid,
  fetchOrCreatePmoByIntegration,
  savePmoItem,
  loadFieldPermissionMap,
  loadGeneralRoleUsers,
  loadTenderTeamUsersFromIntegration,
  canUserEditField, 
  FieldPermissionMap
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
const ACCENT_SOFT = '#e6f0ff';

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
}

const FormApp: React.FC<FormAppProps> = ({
  context,
  pmoListTitle,
  pmoIntegrationLookupName = 'Integration',
  stepsConfig = stepsConfigJson
}) => {
  const sp: SPFI = useMemo(() => getSP(context), [context]);

  const [me, setMe] = useState<{ Email?: string; Title?: string; LoginName?: string } | null>(null);

  // DD של Integration
  const [integrationChoices, setIntegrationChoices] = useState<IDropdownOption[]>([]);
  const [integrationId, setIntegrationId] = useState<number | null>(null);

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
        RFCorTcRFCasPublishedByNTaToBeFi: msg,
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
    'IntegrationTeamDecisionImplement',
    'RevisionIncludesChangeInTenderDo',
    'RFCResponseLetterNo',
    'RFCresponseAsPublishedToBeFilled',
    'Addendum',
    'TenderCommitteeApprovalDate',
    'StatusOfRFCresponseOrTcRFC',
    'DecisionDate',
    'RevisedWordingFinalForPublicatio',
    'RFCorTcRFCasPublishedByNTaToBeFi',
    'DueDateCalculated',
    'ActualDate'
  ];

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

    const filteredIntegrationChoices = useMemo(() => {
    const q = integrationSearch.trim().toLowerCase();
    if (!q) return integrationChoices;

    return integrationChoices.filter(opt =>
      String(opt.text || '').toLowerCase().indexOf(q)!= -1
    );
  }, [integrationSearch, integrationChoices]);

  const loadIntegrationChoices = async () => {
    const items: any[] = await sp.web.lists
      .getById(INTEGRATION_LIST_ID)
      .items.select('Id', INTEGRATION_PREVIEW_FIELD_INTERNAL, 'NTA_x2019_s_x0020_reference')
      .orderBy('Id', false)
      .top(200)();

    const opts: IDropdownOption[] = [];
    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      const id = it.Id as number;
      const raw = it[INTEGRATION_PREVIEW_FIELD_INTERNAL];
      const plain = htmlToPlainText(raw);
      const oneLine = plain.replace(/\r?\n/g, ' ').trim();
      const preview = oneLine.length > 30 ? oneLine.substring(0, 30) + '…' : oneLine;
      const ntaRef = String(it.NTA_x2019_s_x0020_reference || '').trim();
      const textParts: string[] = [];
      if (ntaRef) textParts.push(ntaRef);
        textParts.push(preview ? `${id} — ${preview}` : String(id));
      opts.push({ key: id, text: preview ? `${id} — ${preview}` : String(id) });
    }

    setIntegrationChoices(opts);
    if (!integrationId && opts.length) setIntegrationId(opts[0].key as number);
  };

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

    // ברירת מחדל – למשל All Items או מה שתרצי
    return 'test';
  };


    // טוען את רשימת העמודות מה-View המתאים (לפי TenderPhase)
  const loadViewFieldOrderForPhase = async (tenderPhaseRaw: string) => {
    try {
      const viewName = getViewNameForTenderPhase(tenderPhaseRaw);
      setCurrentViewName(viewName);
      console.log("currentViewName ", currentViewName);

      // לוקחים את ה-View מהרשימה של ה-PMO
      const viewFieldsResp: any = await sp.web.lists
        .getByTitle(pmoListTitle)
        .views.getByTitle(viewName)
        .select('ViewFields')();

      // ב-PnP בדרך כלל השמות הפנימיים נמצאים ב-Items או במערך עצמו
      const internalNames: string[] = (viewFieldsResp?.Items || viewFieldsResp) as string[];

      setViewFieldOrder(internalNames || []);

      // אם את רוצה שכל טאב ב-Pivot יהיה בעצם "ה-View עצמו":
      setStepsInternal({
        [viewName]: internalNames || []
      });
      setActiveStep(viewName);
    } catch (e) {
      console.error('Error loading view fields for phase', e);
      // במקרה של שגיאה – נשאיר את stepsInternal כמו שהיה, כדי שלא ישבר
    }
  };

/*
  const canEdit = (internal: string) => {
  const meEmail = (me?.Email || '').toLowerCase();
  if (!meEmail) return false; // ללא משתמש — קריאה בלבד
  return canUserEditField(
    meEmail,
    internal,
    fieldPermMap,
    roleUsers,
    getOriginatingLineManager()   // 🔹 מעביר M1/M2/M3
  );
};*/


  //🎀
  // האם להציג את RFCResponseLetterNo? רק אם RevisionIncludesChangeInTenderDo = 'Y'
  const showRFCResponseLetterNo = React.useMemo(() => {
    const v = String(pmoDraft?.RevisionIncludesChangeInTenderDo ?? '').trim().toUpperCase();
    return v === 'Y';
  }, [pmoDraft?.RevisionIncludesChangeInTenderDo]);

  // אופציונלי: אם השדה מוסתר – ננקה את הערך כדי שלא יישמר בטעות
  React.useEffect(() => {
    if (!showRFCResponseLetterNo && pmoDraft?.RFCResponseLetterNo) {
      //setPmoDraft((prev: any) => ({ ...prev, RFCResponseLetterNo: null }));
    }
  }, [showRFCResponseLetterNo, pmoDraft?.RFCResponseLetterNo]);
  //🎀

  const loadFormForIntegration = async () => {
    if (!integrationId) return;
    // Integration
    const integ = await fetchIntegrationItemByGuid(sp, integrationId);
    setIntegrationItem(integ);


    // 🔹 אחרי שקיבלנו את פריט ה-Integration – נטען שדות לפי View
    const tenderPhaseRaw = String(integ?.TenderPhase || '');
    await loadViewFieldOrderForPhase(tenderPhaseRaw);
    await loadIntegrationViewFieldOrderForPhase(tenderPhaseRaw);


    // PMO לפי lookup ל-Integration
    const { item: pmoFound, isNew } = await fetchOrCreatePmoByIntegration(
      sp, pmoListTitle, integrationId, pmoIntegrationLookupName
    );
    setPmoItem(pmoFound);
    setPmoDraft(pmoFound);
    // טעינת הרשאות לפי fieldPermission + קבוצות גלובליות + צוות מכרז מתוך ה-Integration
    const [permMap, general, tenderTeamUsers] = await Promise.all([
      loadFieldPermissionMap(sp),
      loadGeneralRoleUsers(sp),
      null//loadTenderTeamUsersFromIntegration(sp, integrationId!)
    ]);


    setFieldPermMap(permMap);
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

    if (isNew) setMsg({ type: MessageBarType.success, text: 'A new PMO item linked to Integration has been created.' });
    else setMsg(null);
    // Labels + Meta ל-PMO
    const pmoMaps = await getFieldMapsByTitle(sp, pmoListTitle);
    setPmoLabels(pmoMaps.internalToTitle);
    const fieldMap = await getFieldInfoMap(sp, pmoListTitle);
    setPmoFieldInfoMap(fieldMap);
    // steps → internal
    const normalized = normalizeStepsConfigToInternal(stepsConfig, pmoMaps.titleToInternal);
    setStepsInternal(normalized);

    const firstStep = Object.keys(normalized)[0] || '';

    setActiveStep(firstStep);

    // PMO צוות מכרז מתוך פריט ה-Integration
    const tenderTeam = await loadTenderTeamUsersFromIntegration(sp, integrationId);

    setRoleUsers(prev => ({ ...prev, PMOTenderTeam: tenderTeam }));

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

  useEffect(() => {
    (async () => {
      try {
        setBusy(true);
        await Promise.all([
          loadIntegrationChoices(),
          loadIntegrationMeta(),
          loadIntegrationLabels(),
          (async () => {
            // הרשאות לפי fieldPermission
            const fp = await loadFieldPermissionMap(sp);
            setFieldPermMap(fp);
            // קבוצות כלליות מ-GeneralRoleDefinition
            const roles = await loadGeneralRoleUsers(sp);
            setRoleUsers(prev => ({
              ...prev,
              FinancialAdvisor: roles.FinancialAdvisor.map(e => e.toLowerCase()),
              Lawyer: roles.Lawyer.map(e => e.toLowerCase()),
              PMOIntegrationTeam: roles.PMOIntegrationTeam.map(e => e.toLowerCase()),
            }));
          })(),
        ]);
      } catch (e: any) {
        setMsg({ type: MessageBarType.error, text: 'Error loading initial data: ' + (e?.message || e) });
      } finally {
        setBusy(false);
      }
    })();
  }, [sp]);

  useEffect(() => {
    (async () => {
      if (!integrationId) return;
      try {
        setBusy(true);
        await loadFormForIntegration();
      } catch (e: any) {
        setMsg({ type: MessageBarType.error, text: 'Error loading the form: ' + (e?.message || e) });
      } finally {
        setBusy(false);
      }
    })();
  }, [integrationId]);

  const onChangeField = (internal: string, value: any) => {
    setPmoDraft((prev: any) => ({ ...prev, [internal]: value }));

    // אם השדה הפך ללא ריק – ננקה שגיאת "שדה חובה"
    setValidationErrors((prev) => {
      if (!prev[internal]) return prev;
      const copy = { ...prev };
      delete copy[internal];
      return copy;
    });
  };

  const onSave = async () => {
    try {
      setBusy(true);
      const saved = await savePmoItem(sp, pmoListTitle, pmoItem.Id, pmoDraft);
      setPmoItem(saved);
      setPmoDraft(saved);
      setMsg({ type: MessageBarType.success, text: 'Saved successfully.' });
      setValidationErrors({});
    } catch (e: any) {
      setMsg({ type: MessageBarType.error, text: 'Save failed: ' + (e?.message || e) });
    } finally {
      setBusy(false);
    }
    if (!pmoItem || !pmoItem.Id) return;
    // ---- ולידציית שדות חובה לפני השמירה ----
    const newErrors: Record<string, string> = {};

    for (let i = 0; i < REQUIRED_FIELDS.length; i++) {
      const internal = REQUIRED_FIELDS[i];

      // אם אין למשתמש הרשאה לערוך את השדה – הוא לא שדה חובה עבורו
      if (!canEditField(internal)) {
        continue;
      }

      if (!isFieldVisibleNow(internal)) {
        continue;
      }

      const v = pmoDraft ? pmoDraft[internal] : undefined;

      let isEmpty = false;
      if (v === null || v === undefined) {
        isEmpty = true;
      } else if (typeof v === 'string') {
        isEmpty = v.trim() === '';
      } else if (Array.isArray(v)) {
        isEmpty = v.length === 0;
      }

      if (isEmpty) {
        newErrors[internal] = 'Required field';
      }
    }

    if (Object.keys(newErrors).length > 0) { 
      setValidationErrors(newErrors);
      setMsg({
        type: MessageBarType.error,
        text: 'Please fill in all required fields marked in red before saving.'
      });
      return; // לא ממשיכים לשמירה ב־SharePoint
    }
    // ---- אם אין שגיאות חובה, ממשיכים לשמור ----
    try {
      setBusy(true);
      const saved = await savePmoItem(sp, pmoListTitle, pmoItem.Id, pmoDraft);
      setPmoItem(saved);
      setPmoDraft(saved);
      setMsg({ type: MessageBarType.success, text: 'Saved successfully.' });
      setValidationErrors({});
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
    ? integrationViewFieldOrder.filter(k => allKeys.indexOf(k)!= -1)
    : allKeys
  );

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
    /*{keys.map(k => {
          const raw = integrationItem[k];
          const info = integrationFieldInfoMap[k];
          let text: string;
          if (info && String(info.TypeAsString || '').toLowerCase().indexOf('date') > -1 && raw) text = formatDateDDMMYYYY(raw);
           text = typeof raw === 'string' && /<[^>]+>/.test(raw)
            ? htmlToPlainText(raw)
            : String(raw ?? '');
          
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
        })} */

    return (
      <Stack tokens={{ childrenGap: 10 }}>
        {keys.map(k => {
  const raw = integrationItem[k];
  const info = integrationFieldInfoMap[k];

  let text: string;

  // 🎯 1. אם זה שדה תאריך (כולל ApplicationDate) – פורמט יפה
  if (info && String(info.TypeAsString || '').toLowerCase().indexOf('date') > -1 && raw) {
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
  const canEdit_ITDI = (appliesVal === 'All Infra 1 tenders' || appliesVal === 'Specific tenders'||
  appliesVal.indexOf('All Infra 1 tenders') !== -1 ||
  appliesVal.indexOf('Specific tenders') !== -1);
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
      tenderPStr.indexOf('phase 2') !== -1; // או התאמה יותר מדויקת אם צריך

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
  /* בניית אפשרויות (keys שמורים נקיים; text עם תאריך)
  const itdiOptions = canEdit_ITDI
    ? [
      { key: 'Done', text: dateText ? `Done — ${dateText}` : 'Done' },
      { key: 'Pending', text: dateText ? `Pending — ${dateText}` : 'Pending' }
    ]
    : [
      { key: 'Not required', text: 'Not required' },
    ];*/

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


/*
  // --- Dynamic hide for RFCResponseLetterNo based on RevisionIncludesChangeInTenderDo ---
  const dynamicHideFields = React.useMemo<string[]>(() => {
    const base = ['Integration', 'IntegrationId', 'Id', 'ID', 'Title'];

     // 1. יישור לפי ה-View: כל שדה שלא מופיע ב-View → מוסתר
    if (viewFieldOrder && viewFieldOrder.length && pmoDraft) {
      const allFields = Object.keys(pmoDraft);
      for (let i = 0; i < allFields.length; i++) {
        const f = allFields[i];
        if (base.indexOf(f) === -1 && viewFieldOrder.indexOf(f) === -1) {
          base.push(f);
        }
      }
    }

    const rv = pmoDraft?.RevisionIncludesChangeInTenderDo;
    const isYes = (() => {
      if (typeof rv === 'boolean') return rv;
      const s = String(rv ?? '').trim().toLowerCase();
      return s === 'y' || s === 'yes' || s === 'true' || s === '1';
    })();

    if (!isYes) base.push('RFCResponseLetterNo');

    return base;
  }, [pmoDraft, pmoDraft?.RevisionIncludesChangeInTenderDo, viewFieldOrder]);
*/
const dynamicHideFields = React.useMemo<string[]>(() => {
  const base = ['Integration', 'IntegrationId', 'Id', 'ID', 'Title'];

  // 1. יישור לפי ה-View: כל שדה שלא מופיע ב-View → מוסתר
  if (viewFieldOrder && viewFieldOrder.length && pmoDraft) {
    const allFields = Object.keys(pmoDraft);
    for (let i = 0; i < allFields.length; i++) {
      const f = allFields[i];
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

    if (internal === 'RFCresponseAsPublishedToBeFilled' || internal === 'StatusOfRFCresponseOrTcRFC') {
      if (decision !== 'Accept') return false;
    }

    if (internal === 'Addendum' || internal === 'TenderCommitteeApprovalDate') {
      const isNo = (revInc === 'n' || revInc === 'no' || revInc === 'false' || revInc === '0');
      if (isNo) return false;
    }

    if (internal === 'RFCorTcRFCasPublishedByNTaToBeFi') {
      if (tenderPhaseStr === 'phase 1 – preparation of tender documents') return false;
    }

    if (internal === 'RevisedWordingFinalForPublicatio') {
      if (tenderPhaseStr !== 'phase 1 – preparation of tender documents') return false;
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

  /*
  const renderPmoEditableBySteps = () => {
    const stepNames = Object.keys(stepsInternal);
    if (!stepNames.length) {
      return (
        <EditableFields
          item={pmoDraft}
          onChange={onChangeField}
          hideFields={dynamicHideFields}
          internalToTitle={pmoLabels}
          fieldInfoMap={pmoFieldInfoMap}
          canEdit={canEdit}
          placeholderMap={placeholders}
          choiceOverrides={{
            [TARGET_FIELD]: itdiOptions
          }}
          tenderPhase={String(integrationItem?.TenderPhase || '')}
          requiredMap={requiredMap}
          errorMap={validationErrors}
        />
      );
    }
    return (
      <Pivot
        selectedKey={activeStep}
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
                fieldOrder={stepsInternal[step] || []}
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
              />
            </div>
          </PivotItem>
        ))}
      </Pivot>
    );
  };
  */

  const renderPmoEditableBySteps = () => {
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
      />
    );
  }
  const getStepFieldOrder = (step: string): string[] => {
      const stepFields = stepsInternal[step] || [];

      // אם יש View – נכבד אותו קודם, ואז נסנן לפי ה-step
      if (viewFieldOrder && viewFieldOrder.length) {
        if (stepFields.length) {
          return viewFieldOrder.filter(f => stepFields.indexOf(f) !== -1);
        }
        return viewFieldOrder;
      }

      // בלי View – נשאר רק עם ה-step המקורי
      return stepFields;
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
            />
          </div>
        </PivotItem>
      ))}
    </Pivot>
  );
};


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
                Integration item selection
              </div>

              {/* 🔍 שורת חיפוש + דרופדאון */}
              <Stack horizontal tokens={{ childrenGap: 8 }}>
                <TextField
                  label="Filter by NTA reference"
                  value={integrationSearch}
                  onChange={(_, v) => setIntegrationSearch(v ?? '')}
                  styles={{ root: { minWidth: 220 } }}
                />

                <Dropdown
                  label="Search by NTA Reference"
                  placeholder={
                    busy
                      ? 'Loading…'
                      : (filteredIntegrationChoices.length ? 'Select item' : 'No items')
                  }
                  options={filteredIntegrationChoices}
                  selectedKey={integrationId ?? undefined}
                  disabled={busy || !filteredIntegrationChoices.length}
                  onChange={(_, opt) => setIntegrationId(opt ? (opt.key as number) : null)}
                  styles={{
                    dropdown: { minWidth: 320 },
                    label: { fontWeight: 600 }
                  }}
                />
              </Stack>
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
          </Stack>
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
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text={busy ? 'Saveing' : 'Save'}
              onClick={onSave}
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
          </Stack>
        </Stack>
      </Stack>
    </div>
  );
};

export default FormApp;



/*
// src/webparts/smartForm/components/FormApp.tsx
// src/webparts/smartForm/components/FormApp.tsx
import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import {
  Stack, Pivot, PivotItem,
  PrimaryButton, DefaultButton,
  MessageBar, MessageBarType,
  Dropdown, IDropdownOption
} from '@fluentui/react';
import type { WebPartContext } from '@microsoft/sp-webpart-base';
import type { SPFI } from '@pnp/sp';
import { getSP } from './pnpjsConfig';
import { isSystemField } from '../shared/constants';
import { INTEGRATION_LIST_ID, INTEGRATION_PREVIEW_FIELD_INTERNAL } from '../shared/internalNames';
import {
  getFieldMapsByTitle,
  getFieldInfoMap,
  fetchIntegrationItemByGuid,
  fetchOrCreatePmoByIntegration,
  savePmoItem,
  loadFieldPermissionMap,
  loadGeneralRoleUsers,
  loadTenderTeamUsersFromIntegration,
  canUserEditField
} from '../shared/data';
import stepsConfigJson from '../stepsConfig.json';
import '@pnp/sp/site-users/web';
import EditableFields from './EditableFields';

export interface FormAppProps {
  context: WebPartContext;
  pmoListTitle: string;
  pmoIntegrationLookupName?: string;
  stepsConfig?: Record<string, string[]>;
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

const FormApp: React.FC<FormAppProps> = ({
  context,
  pmoListTitle,
  pmoIntegrationLookupName = 'Integration',
  stepsConfig = stepsConfigJson
}) => {
  
  //const DECISION_FIELD = 'DecisionRegardingProposedChange'; // התאימי לשם האמיתי
  //const RFC_FIELD = 'RFCresponseAsPublishedToBeFilled';


  const sp: SPFI = useMemo(() => getSP(context), [context]);

  const [me, setMe] = useState<{ Email?: string; Title?: string; LoginName?: string } | null>(null);
  
  // DD של Integration
  const [integrationChoices, setIntegrationChoices] = useState<IDropdownOption[]>([]);
  const [integrationId, setIntegrationId] = useState<number | null>(null);

  // Integration (תצוגה בלבד) + תוויות + מטא
  const [integrationItem, setIntegrationItem] = useState<any>(null);
  const [integrationFieldInfoMap, setIntegrationFieldInfoMap] = useState<Record<string, any>>({});
  const [integrationLabels, setIntegrationLabels] = useState<Record<string, string>>({});

  // PMO (עריכה) + תוויות + מטא
  const [pmoItem, setPmoItem] = useState<any>(null);
  const [pmoDraft, setPmoDraft] = useState<any>({});
  // placeholder לשדות ספציפיים כשיש Accept
  const placeholders = React.useMemo<Partial<Record<string, string>>>(() => {
    const decision = String(pmoDraft?.DecisionRegardingProposedChange || '').trim();
    if (decision === 'Accept') {
      const msg = 'להזין כאן ניסוח סופי לפרסום';
      return {
        RevisedWordingFinalForPublicatio: msg,
        RFCorTcRFCasPublishedByNTaToBeFi: msg,
      };
    }
    return {};
  }, [pmoDraft?.DecisionRegardingProposedChange]);

  const [pmoLabels, setPmoLabels] = useState<Record<string, string>>({});
  const [pmoFieldInfoMap, setPmoFieldInfoMap] = useState<Record<string, any>>({});

  // שלבים
  const [stepsInternal, setStepsInternal] = useState<Record<string, string[]>>({});
  const [activeStep, setActiveStep] = useState<string>('');

  // הרשאות שדה
  const [fieldPermMap, setFieldPermMap] = useState<Record<string, string[]>>({});
  const [roleUsers, setRoleUsers] = useState<{
    FinancialAdvisor: string[];
    Lawyer: string[];
    PMOIntegrationTeam: string[];
    PMOTenderTeam: string[];
  }>({
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
    'IntegrationTeamDecisionImplement',
    'RevisionIncludesChangeInTenderDo',
    'RFCResponseLetterNo',
    'RFCresponseAsPublishedToBeFilled',
    'Addendum',
    'TenderCommitteeApprovalDate',
    'StatusOfRFCresponseOrTcRFC',
    'DecisionDate',
    'RevisedWordingFinalForPublicatio',
    'RFCorTcRFCasPublishedByNTaToBeFi',
    'DueDateCalculated',
    'ActualDate'
  ];

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


  const loadIntegrationChoices = async () => {
    const items: any[] = await sp.web.lists
      .getById(INTEGRATION_LIST_ID)
      .items.select('Id', INTEGRATION_PREVIEW_FIELD_INTERNAL)
      .orderBy('Id', false)
      .top(200)();

    const opts: IDropdownOption[] = [];
    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      const id = it.Id as number;
      const raw = it[INTEGRATION_PREVIEW_FIELD_INTERNAL];
      const plain = htmlToPlainText(raw);
      const oneLine = plain.replace(/\r?\n/g, ' ').trim();
      const preview = oneLine.length > 30 ? oneLine.substring(0, 30) + '…' : oneLine;
      opts.push({ key: id, text: preview ? `${id} — ${preview}` : String(id) });
    }

    setIntegrationChoices(opts);
    if (!integrationId && opts.length) setIntegrationId(opts[0].key as number);
  };

  const loadIntegrationMeta = async () => {
    const fields = await sp.web.lists
      .getById(INTEGRATION_LIST_ID)
      .fields
      .select('Title','InternalName','Hidden','ReadOnlyField','TypeAsString')();

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
    const fields = await sp.web.lists.getById(INTEGRATION_LIST_ID).fields.select('Title','InternalName')();
    const map: Record<string, string> = {};
    for (let i = 0; i < fields.length; i++) {
      const f: any = fields[i];
      map[f.InternalName] = f.Title || f.InternalName;
    }
    setIntegrationLabels(map);
  };

  const canEdit = (internal: string) => {
    const meEmail = (me?.Email || '').toLowerCase();
    if (!meEmail) return false;           // ללא משתמש — קריאה בלבד
    return canUserEditField(meEmail, internal, fieldPermMap, roleUsers);
    };

    //🎀
    // האם להציג את RFCResponseLetterNo? רק אם RevisionIncludesChangeInTenderDo = 'Y'
    const showRFCResponseLetterNo = React.useMemo(() => {
      const v = String(pmoDraft?.RevisionIncludesChangeInTenderDo ?? '').trim().toUpperCase();
      return v === 'Y';
    }, [pmoDraft?.RevisionIncludesChangeInTenderDo]);

    // רשימת שדות להסתיר קבוע + הסתרה דינמית של RFCResponseLetterNo
    
    // אופציונלי: אם השדה מוסתר – ננקה את הערך כדי שלא יישמר בטעות
    React.useEffect(() => {
      if (!showRFCResponseLetterNo && pmoDraft?.RFCResponseLetterNo) {
        setPmoDraft((prev: any) => ({ ...prev, RFCResponseLetterNo: null }));
      }
    }, [showRFCResponseLetterNo]);
    //🎀
  
  const loadFormForIntegration = async () => {
    if (!integrationId) return;
    console.log("a. 😏");
    // Integration
    const integ = await fetchIntegrationItemByGuid(sp, integrationId);
    console.log("b. 😏");
    setIntegrationItem(integ);
    console.log("c. 😏");
    // PMO לפי lookup ל-Integration
    const { item: pmoFound, isNew } = await fetchOrCreatePmoByIntegration(
      sp, pmoListTitle, integrationId, pmoIntegrationLookupName
    );
    console.log("d. 😏");
    setPmoItem(pmoFound);
    console.log("e. 😏");
    setPmoDraft(pmoFound);
    // טעינת הרשאות לפי fieldPermission + קבוצות גלובליות + צוות מכרז מתוך ה-Integration
    const [permMap, general, tenderTeamUsers] = await Promise.all([
      loadFieldPermissionMap(sp),
      loadGeneralRoleUsers(sp),
      loadTenderTeamUsersFromIntegration(sp, integrationId!)
    ]);

    console.log("🥰 permMap ", permMap ," general ", general, " tenderTeamUsers ", tenderTeamUsers);

    setFieldPermMap(permMap);
    setRoleUsers({
      FinancialAdvisor: general.FinancialAdvisor,
      Lawyer: general.Lawyer,
      PMOIntegrationTeam: general.PMOIntegrationTeam,
      PMOTenderTeam: tenderTeamUsers,   // ← חשוב: כאן נכנסת הרשאת "PMO צוות מכרז"
    });

    console.log("f. 😏");
    if (isNew) setMsg({ type: MessageBarType.success, text: 'נוצר פריט PMO חדש מקושר ל-Integration.' });
    else setMsg(null);
    console.log("g. 😏");
    // Labels + Meta ל-PMO
    const pmoMaps = await getFieldMapsByTitle(sp, pmoListTitle);
    console.log("h. 😏");
    setPmoLabels(pmoMaps.internalToTitle);
    console.log("i. 😏");
    const fieldMap = await getFieldInfoMap(sp, pmoListTitle);
    console.log("j. 😏");
    setPmoFieldInfoMap(fieldMap);
    console.log("k. 😏");
    // steps → internal
    const normalized = normalizeStepsConfigToInternal(stepsConfig, pmoMaps.titleToInternal);
    console.log("l. 😏");
    setStepsInternal(normalized);
    console.log("m. 😏");
    const firstStep = Object.keys(normalized)[0] || '';
    console.log("n. 😏");
    setActiveStep(firstStep);
    console.log("o. 😏");
    // PMO צוות מכרז מתוך פריט ה-Integration
    const tenderTeam = await loadTenderTeamUsersFromIntegration(sp, integrationId);
    console.log("p. 😏");
    setRoleUsers(prev => ({ ...prev, PMOTenderTeam: tenderTeam }));
    console.log("q. 😏");
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

  useEffect(() => {
    (async () => {
      try {
        setBusy(true);
        await Promise.all([
          loadIntegrationChoices(),
          loadIntegrationMeta(),
          loadIntegrationLabels(),
          (async () => {
            // הרשאות לפי fieldPermission
            const fp = await loadFieldPermissionMap(sp);
            setFieldPermMap(fp);
            // קבוצות כלליות מ-GeneralRoleDefinition
            const roles = await loadGeneralRoleUsers(sp);
            setRoleUsers(prev => ({
              ...prev,
              FinancialAdvisor: roles.FinancialAdvisor.map(e => e.toLowerCase()),
              Lawyer: roles.Lawyer.map(e => e.toLowerCase()),
              PMOIntegrationTeam: roles.PMOIntegrationTeam.map(e => e.toLowerCase()),
            }));
          })(),
        ]);
      } catch (e: any) {
        setMsg({ type: MessageBarType.error, text: 'שגיאה בטעינת נתונים ראשוניים: ' + (e?.message || e) });
      } finally {
        setBusy(false);
      }
    })();
  }, [sp]);

  useEffect(() => {
    (async () => {
      if (!integrationId) return;
      try {
        console.log("1 🦜");
        setBusy(true);
        console.log("2 🦜");
        await loadFormForIntegration();
        console.log("3 🦜");

      } catch (e: any) {
        setMsg({ type: MessageBarType.error, text: 'שגיאה בטעינת הטופס: ' + (e?.message || e) });
      } finally {
        console.log("🤖 1");
        setBusy(false);
        console.log("🤖 2");
      }
    })();
  }, [integrationId]);

  
   const onChangeField = (internal: string, value: any) => {
    setPmoDraft((prev: any) => ({ ...prev, [internal]: value }));

    // אם השדה הפך ללא ריק – ננקה שגיאת "שדה חובה"
    setValidationErrors((prev) => {
      if (!prev[internal]) return prev;
      const copy = { ...prev };
      delete copy[internal];
      return copy;
    });
  };

  

    const onSave = async () => {
    if (!pmoItem || !pmoItem.Id) return;
      console.log("🧋in on save ");
    // ---- ולידציית שדות חובה לפני השמירה ----
    const newErrors: Record<string, string> = {};

    for (let i = 0; i < REQUIRED_FIELDS.length; i++) {
      const internal = REQUIRED_FIELDS[i];

      // אם אין למשתמש הרשאה לערוך את השדה – הוא לא שדה חובה עבורו
      if (!canEditField(internal)){console.log("🐕‍🦺this field is readonly ", internal);
        continue;
        
      } 

      if (!isFieldVisibleNow(internal)){
        console.log("🦪 this field is not diplayed ", internal);
        continue;

      } 
      console.log("✒️this is an editable field internal ", internal);

      const v = pmoDraft ? pmoDraft[internal] : undefined;

      let isEmpty = false;
      if (v === null || v === undefined) {
        isEmpty = true;
      } else if (typeof v === 'string') {
        isEmpty = v.trim() === '';
      } else if (Array.isArray(v)) {
        isEmpty = v.length === 0;
      }

      if (isEmpty) {
        newErrors[internal] = 'שדה חובה';
      }
    }

    if (Object.keys(newErrors).length > 0) {
      console.log("🍥Object.keys(newErrors).length ", Object.keys(newErrors).length, "Object.keys(newErrors)", Object.keys(newErrors));
      setValidationErrors(newErrors);
      setMsg({
        type: MessageBarType.error,
        text: 'נא למלא את כל השדות החובה המסומנים באדום לפני שמירה.'
      });
      return; // לא ממשיכים לשמירה ב־SharePoint
    }
    // ---- אם אין שגיאות חובה, ממשיכים לשמור ----
    try {
      setBusy(true);
      const saved = await savePmoItem(sp, pmoListTitle, pmoItem.Id, pmoDraft);
      setPmoItem(saved);
      setPmoDraft(saved);
      setMsg({ type: MessageBarType.success, text: 'נשמר בהצלחה.' });
      setValidationErrors({});
    } catch (e: any) {
      setMsg({ type: MessageBarType.error, text: 'שמירה נכשלה: ' + (e?.message || e) });
    } finally {
      setBusy(false);
    }
  };
  

  

  const renderIntegrationReadonly = () => {
    if (!integrationItem) return null;

    const keys = Object.keys(integrationItem).filter(k => {
      if (isSystemField(k)) return false;
      if (k === 'Id' || k === 'ID' || k === 'Title') return false;
      const info = integrationFieldInfoMap[k];
      if (info && info.Hidden === true) return false;
      return true;
    });

    return (
      <Stack tokens={{ childrenGap: 6 }}>
        {keys.map(k => {
          const raw = integrationItem[k];
          const text = typeof raw === 'string' && /<[^>]+>/.test(raw)
            ? htmlToPlainText(raw)
            : String(raw ?? '');
          return (
            <div key={k} style={{ padding: 6, background: '#f9fafb', borderRadius: 8, border: '1px solid #eee' }}>
              <div style={{ fontSize: 12, color: '#6b7280' }}>
                {integrationLabels[k] || k}
              </div>
              <div style={{ fontWeight: 600, whiteSpace: 'pre-wrap' }}>{text}</div>
            </div>
          );
        })}
      </Stack>
    );
  };

  const canEditField = (internal: string): boolean => {
  const emailLc = String(me?.Email || '').toLowerCase();
  if (!emailLc) return false;

  // כלל מיוחד לשדה IntegrationTeamDecisionImplement
  if (internal === TARGET_FIELD) {
    // ניתן לעריכה רק אם Applies הוא All Infra 1 tenders / Specific tenders
    return canEdit_ITDI && canUserEditField(emailLc, internal, fieldPermMap, {
      FinancialAdvisor: roleUsers.FinancialAdvisor,
      Lawyer: roleUsers.Lawyer,
      PMOIntegrationTeam: roleUsers.PMOIntegrationTeam,
      PMOTenderTeam: roleUsers.PMOTenderTeam
    });
  }

  // אחרת – ברירת המחדל (לפי fieldPermission)
  return canUserEditField(emailLc, internal, fieldPermMap, roleUsers);
};



  // --- לוגיקת תלות לשדה IntegrationTeamDecisionImplement ---
  const APPLIES_FIELD = 'DecisionAppliesToOtherWorksTende';
  const TARGET_FIELD  = 'IntegrationTeamDecisionImplement';
  const DATE_FIELD    = 'DateForIntegrationTeamDecisionIm';

  // האם "Applies" מאפשר עריכה?
  const appliesVal = String(pmoDraft?.[APPLIES_FIELD] || '').trim();
  const canEdit_ITDI = (appliesVal === 'All Infra 1 tenders' || appliesVal === 'Specific tenders');

  // הפקת תאריך לתצוגה
  let dateText = '';
  try {
    const rawDate = pmoDraft?.[DATE_FIELD];
    if (rawDate) {
      const d = new Date(rawDate);
      if (!isNaN(d.getTime())) {
        // תצוגה נוחה; אפשר להחליף ל-toLocaleDateString('he-IL', {year:'numeric', month:'2-digit', day:'2-digit'})
        dateText = d.toLocaleDateString('he-IL');
      }
    }
  } catch { }

  // בניית אפשרויות (keys שמורים נקיים; text עם תאריך)
  const itdiOptions = canEdit_ITDI
    ? [
        { key: 'Done',        text: dateText ? `Done — ${dateText}` : 'Done' },
        { key: 'Pending',     text: dateText ? `Pending — ${dateText}` : 'Pending' },
        { key: 'Not required', text: 'Not required' },
      ]
    : [
        { key: 'Not required', text: 'Not required' },
      ];

  // אם לא ניתן לערוך – נכפה ערך "Not required"
  useEffect(() => {
    if (!canEdit_ITDI) {
      if (pmoDraft?.[TARGET_FIELD] !== 'Not required') {
        setPmoDraft((prev: any) => ({ ...(prev || {}), [TARGET_FIELD]: 'Not required' }));
      }
    }
  }, [canEdit_ITDI, pmoDraft?.[TARGET_FIELD]]);




    // --- Dynamic hide for RFCResponseLetterNo based on RevisionIncludesChangeInTenderDo ---
  const dynamicHideFields = React.useMemo<string[]>(() => {
    // בסיס ההסתרה – כמו שהיה עד עכשיו
    const base = ['Integration', 'IntegrationId', 'Id', 'ID', 'Title'];

    // נזהה את ערך השדה: יכול להיות בוליאני (true/false) או טקסט (Y/N/Yes/No/True/False/1/0)
    const rv = pmoDraft?.RevisionIncludesChangeInTenderDo;
    const isYes = (() => {
      if (typeof rv === 'boolean') return rv;
      const s = String(rv ?? '').trim().toLowerCase();
      return s === 'y' || s === 'yes' || s === 'true' || s === '1';
    })();

    // אם זה לא Yes → מסתירים את RFCResponseLetterNo
    if (!isYes) base.push('RFCResponseLetterNo');

    return base;
  }, [pmoDraft?.RevisionIncludesChangeInTenderDo]);

  // האם השדה הזה *מוצג כרגע* בטופס (לפי אותם תנאים של EditableFields)
  const isFieldVisibleNow = (internal: string): boolean => {
    // מוסתרים קבוע
    if (['Integration', 'IntegrationId', 'Id', 'ID', 'Title'].indexOf(internal) > -1) {
      return false;
    }

    // מוסתרים דינאמית לפי dynamicHideFields
    if (dynamicHideFields.indexOf(internal) > -1) {
      return false;
    }

    // ---- תנאים כמו ב-EditableFields ----

    const decision = String(pmoDraft?.DecisionRegardingProposedChange || '').trim();
    const revInc = String(pmoDraft?.RevisionIncludesChangeInTenderDo || '').trim().toLowerCase();
    const tenderPhaseStr = String(integrationItem?.TenderPhase || '').trim().toLowerCase();

    // שדות שתלויים ב-DecisionRegardingProposedChange = Accept
    if (internal === 'RFCresponseAsPublishedToBeFilled' || internal === 'StatusOfRFCresponseOrTcRFC') {
      if (decision !== 'Accept') return false;
    }

    // Addendum + TenderCommitteeApprovalDate רק אם RevisionIncludesChangeInTenderDo היא "כן"
    if (internal === 'Addendum' || internal === 'TenderCommitteeApprovalDate') {
      const isNo = (revInc === 'n' || revInc === 'no' || revInc === 'false' || revInc === '0');
      if (isNo) return false;
    }

    // תנאים לפי TenderPhase
    if (internal === 'RFCorTcRFCasPublishedByNTaToBeFi') {
      if (tenderPhaseStr === 'phase 1 – preparation of tender documents') return false;
    }

    if (internal === 'RevisedWordingFinalForPublicatio') {
      if (tenderPhaseStr !== 'phase 1 – preparation of tender documents') return false;
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

    return true;
  };



  const renderPmoEditableBySteps = () => {
    const stepNames = Object.keys(stepsInternal);
    if (!stepNames.length) {
      return (
        <EditableFields
          item={pmoDraft}
          onChange={onChangeField}
          hideFields={dynamicHideFields}
          internalToTitle={pmoLabels}
          fieldInfoMap={pmoFieldInfoMap}
          canEdit={canEdit}   
          placeholderMap={placeholders}
          choiceOverrides={{
            [TARGET_FIELD]: itdiOptions
          }}
          tenderPhase={String(integrationItem?.TenderPhase || '')}  
          requiredMap={requiredMap}
          errorMap={validationErrors}
        />
      );
    }
    return (
      <Pivot selectedKey={activeStep} onLinkClick={(i) => setActiveStep(i?.props.itemKey || '')}>
        {stepNames.map(step => (
          <PivotItem headerText={step} itemKey={step} key={step}>
            <div style={{ marginTop: 10 }}>
              <EditableFields
                item={pmoDraft}
                onChange={onChangeField}
                fieldOrder={stepsInternal[step] || []}
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
              />
            </div>
          </PivotItem>
        ))}
      </Pivot>
    );
  };

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      {}
      {me?.Email && (
        <div style={{ padding: '6px 10px', background: '#eef6ff', border: '1px solid #d0e3ff', borderRadius: 8, marginBottom: 8 }}>
          <strong>Current user:</strong> {me.Title || me.LoginName} &nbsp;|&nbsp; <strong>Email:</strong> {me.Email}
        </div>
      )}

      {msg ? <MessageBar messageBarType={msg.type}>{msg.text}</MessageBar> : null}

      <Stack horizontal tokens={{ childrenGap: 8 }} style={{ alignItems: 'end' }}>
        <Dropdown
          label="בחר/י פריט Integration להצגה/עריכה"
          placeholder={busy ? 'טוען…' : (integrationChoices.length ? 'בחר/י פריט' : 'אין פריטים')}
          options={integrationChoices}
          selectedKey={integrationId ?? undefined}
          disabled={busy || !integrationChoices.length}
          onChange={(_, opt) => setIntegrationId(opt ? (opt.key as number) : null)}
          styles={{ dropdown: { minWidth: 420 } }}
        />
        <DefaultButton text="רענון רשימה" onClick={loadIntegrationChoices} disabled={busy} />
      </Stack>

      <Stack horizontal wrap tokens={{ childrenGap: 24 }} styles={{ root: { alignItems: 'flex-start' } }}>
        <Stack grow>
          <h3 style={{ marginTop: 0 }}>Integration – תצוגה</h3>
          {renderIntegrationReadonly()}
        </Stack>

        <Stack grow>
          <h3 style={{ marginTop: 0 }}>PMO decisions – עריכה</h3>
          {renderPmoEditableBySteps()}
        </Stack>
      </Stack>

      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton text={busy ? 'שומר…' : 'שמירה'} onClick={onSave} disabled={busy || !pmoItem} />
        <DefaultButton text="רענון טופס" onClick={loadFormForIntegration} disabled={busy || !integrationId} />
      </Stack>
    </Stack>
  );
};

export default FormApp;

*/