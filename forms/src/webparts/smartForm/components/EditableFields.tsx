
// src/webparts/smartForm/components/EditableFields.tsx
import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { isSystemField } from '../shared/constants';

const REQUIRED_FIELDS: string[] = [
  'DecisionRegardingProposedChange',
  'DecisionAppliesToOtherWorksTende',
  'RevisionIncludesChangeInTenderDo',
  'DecisionDate',
  'DueDateCalculated',
  'ActualDate'
];



interface Props {
  item: any;
  onChange: (internalName: string, value: any) => void;
  fieldOrder?: string[];
  hideFields?: string[];
  internalToTitle?: Record<string, string>;
  fieldInfoMap?: Record<string, any>;
  // חדש: פונקציה שמחליטה אם שדה מסוים ניתן לעריכה (ברירת מחדל true)
  canEdit?: (internalName: string) => boolean;
  placeholderMap?: Partial<Record<string, string>>;
  choiceOverrides?: Record<string, { key: string; text: string }[]>;
  tenderPhase?: string;
  // חדש: אילו שדות הם חובה
  requiredMap?: Record<string, boolean>;
  // חדש: טקסט שגיאה לשדות (שדה -> טקסט)
  errorMap?: Record<string, string>;
  labelOverrides?: Record<string, string>;
}

// ===== עזרי טקסט =====
const lpad = (s: any, len: number, ch: string = '0') => {
  let t = String(s == null ? '' : s);
  while (t.length < len) t = ch + t;
  return t;
};
export const formatDateDDMMYYYY = (d?: Date | string | null): string => {
  if (!d) return '';
  const dt = typeof d === 'string' ? new Date(d) : d;
  if (!(dt instanceof Date) || isNaN(dt.getTime())) return '';
  const dd = lpad(dt.getDate(), 2);
  const mm = lpad(dt.getMonth() + 1, 2);
  const yyyy = dt.getFullYear();
  return `${dd}/${mm}/${yyyy}`;
};

const strHas = (a?: string, frag?: string) =>
  String(a || '').toLowerCase().indexOf(String(frag || '').toLowerCase()) > -1;

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

// ===== זיהוי סוגי שדות =====
function isMultiField(info?: any): boolean {
  if (!info) return false;
  if (info.AllowMultipleValues === true) return true;
  return strHas(info?.TypeAsString, 'multi');
}
function isLookupOrUser(info?: any): boolean {
  if (!info || !info.TypeAsString) return false;
  const t = String(info.TypeAsString || '').toLowerCase();
  return t.indexOf('lookup') > -1 || t.indexOf('user') > -1;
}
function isChoice(info?: any): boolean {
  if (!info || !info.TypeAsString) return false;
  const t = String(info.TypeAsString || '').toLowerCase();
  return t === 'choice' || t === 'multichoice';
}
function isTaxonomy(info?: any): boolean {
  if (!info || !info.TypeAsString) return false;
  return strHas(info.TypeAsString, 'taxonomy');
}
function isDate(info?: any): boolean {
  if (!info || !info.TypeAsString) return false;
  const t = String(info.TypeAsString || '').toLowerCase();
  return t.indexOf('date') > -1;
}
function isNumber(info?: any): boolean {
  if (!info || !info.TypeAsString) return false;
  const t = String(info.TypeAsString || '').toLowerCase();
  return t === 'number' || t === 'currency';
}
function isBoolean(info?: any): boolean {
  if (!info || !info.TypeAsString) return false;
  return String(info.TypeAsString || '').toLowerCase() === 'boolean';
}
function isRichText(info?: any): boolean {
  if (!info) return false;
  return info.RichText === true || String(info.TypeAsString || '').toLowerCase() === 'note';
}

function isUrl(info?: any): boolean {
  if (!info || !info.TypeAsString) return false;
  const t = String(info.TypeAsString || '').toLowerCase();
  return t === 'url';
}


// ===== טקסט עזרה =====
function helpTextFor(info?: any): string | undefined {
  if (!info) return undefined;
  if (isBoolean(info)) return 'Binary value: yes/no.';
  if (isDate(info)) return 'Date/Date-Time (can be selected from the date range).';
  if (isNumber(info)) return 'Number (digits and decimal point).';
  if (isChoice(info) && isMultiField(info)) return 'Multiple selection from a list of options.';
  if (isChoice(info)) return 'A single selection from a list of options.';
  if (isLookupOrUser(info) && isMultiField(info)) return 'Link/User – Identifiers (Id) separated by comma.';
  if (isLookupOrUser(info)) return 'Link/User – Unique Identifier (Id).';
  if (isRichText(info)) return 'Rich text will be displayed as plain text (without HTML tags).';
  if (isTaxonomy(info)) return 'Taxonomy field.';
  return undefined;
}

// ===== עזר מספרי =====
function toNumberOrNull(s?: string): number | null {
  if (s === undefined || s === null) return null;
  const t = String(s).trim();
  if (!t) return null;
  const normalized = t.replace(',', '.');
  const n = parseFloat(normalized);
  return isFinite(n) ? n : null;
}

// ===== עיצוב בסיסי לכל הקונטרולים =====
const baseControlMaxWidth = 480;
const textStyles = {
  label: { fontWeight: 600 },
  fieldGroup: { borderRadius: 10 },
};
const dropdownStyles = {
  label: { fontWeight: 600 },
  root: { maxWidth: baseControlMaxWidth },
};
const datePickerStyles = {
  root: { maxWidth: baseControlMaxWidth },
  textField: { ...textStyles },
};

// ===== הקומפוננטה =====
const EditableFields: React.FC<Props> = ({
  item,
  onChange,
  fieldOrder = [],
  hideFields = [],
  internalToTitle = {},
  fieldInfoMap = {},
  canEdit = () => false,
  placeholderMap = {},
  choiceOverrides = {},
  tenderPhase,
  requiredMap = {},
  errorMap = {},
  labelOverrides = {},
}) => {
  const extraHide = ['Integration', 'IntegrationId', 'Id', 'ID', 'Title'];
  console.log("🥫 fieldOrder ", fieldOrder, "\n🍼 hideFields", hideFields);

  const keys = React.useMemo(() => {
    console.log("🎋Object.keys(item)", Object.keys(item));
    const all = Object.keys(item || {}).filter(k =>
      hideFields.indexOf(k) === -1 &&
      extraHide.indexOf(k) === -1 &&
      !isSystemField(k)&&
      fieldOrder.indexOf(k) > -1 
    );

    console.log("all- 1 🔑 ", all);

    const ordered: string[] = [];
    const seen: Record<string, boolean> = {};
    for (let i = 0; i < fieldOrder.length; i++) {
      const f = fieldOrder[i];
      if (all.indexOf(f) > -1 && !seen[f]) { ordered.push(f); seen[f] = true; }
    }

    for (let i = 0; i < all.length; i++) {
      const f = all[i];
      if (!seen[f]) { ordered.push(f); seen[f] = true; }
    }
    console.log("all 🔑 ", all);
    console.log("ordered 🔑 ", ordered);

    return ordered;
  }, [item, fieldOrder, hideFields]);

  let displayRevisionIncludesChangeInTenderDo = true;
  return (
    <Stack
      tokens={{ childrenGap: 12 }}
      styles={{
        root: {
          paddingTop: 4,
        }
      }}
    >
      {keys.map((internal) => {
        if (!internal) return null;
        const tenderPhaseStr = String(tenderPhase || '').trim().toLowerCase();
        console.log("🍩🧋 tenderPhaseStr ", tenderPhaseStr);
        const val = item[internal];
        const info = fieldInfoMap[internal];


        console.log("🫏 const label = (labelOverrides && labelOverrides[internal]) || internalToTitle[internal] ||internal; ",
          "label  ", (labelOverrides && labelOverrides[internal]) || internalToTitle[internal] ||internal, "labelOverrides  ", labelOverrides,"  labelOverrides[internal] ",
          labelOverrides[internal],"internalToTitle[internal]",  internalToTitle[internal], "  internal   " ,internal);
        const label = (labelOverrides && labelOverrides[internal])||(internalToTitle?.[internal]) ;

        //|| internalToTitle[internal] || internal;
        const hint = helpTextFor(info);
        const editable = canEdit ? !!canEdit(internal) : true;
        const required = editable && REQUIRED_FIELDS.indexOf(internal) > -1;
        const error = errorMap[internal];

        console.log('🤗tenderPhaseLc ', tenderPhaseStr);
        let placeholder = placeholderMap?.[internal];
        
        if(internal === 'RFCorTcRFCasPublishedByNTaToBeFi'){
          const decision = String(item?.DecisionRegardingProposedChange || '').trim();
          console.log('🍩decision ', decision);

          if (tenderPhaseStr.indexOf('phase 1') > -1) {
            return null;
          }
          if(decision === 'Accept'){
            placeholder = 'Enter final version for publication here';
          }
          
        }

        if(internal === 'Assignedto' || internal === 'SubCategory'){
          if(tenderPhaseStr.indexOf('phase 1') === -1){
            return null;
          }
        }


        if (internal === 'dog') {
            return null;

        }

        if (internal === 'RevisedWordingFinalForPublicatio') {
          if (tenderPhaseStr != 'phase 1 – preparation of tender documents') {
            return null;
          }
        }

        
        if (
          internal === 'RevisionIncludesChangeInTenderDo' ||
          internal === 'RFCResponseLetterNo' ||
          internal === 'RFCresponseAsPublishedToBeFilled'
        ) {
          if (tenderPhaseStr != 'phase 2 - bidders’ requests for clarifications (rfcs) of tender documents' && tenderPhaseStr.indexOf('phase 2') === -1) {
            if(internal === 'RevisionIncludesChangeInTenderDo'){
              displayRevisionIncludesChangeInTenderDo = false;
            }
            return null;
          }
        }

        if (internal === 'RFCresponseAsPublishedToBeFilled') {
          const decision = String(item?.DecisionRegardingProposedChange || '').trim();
          console.log('🍩decision ', decision);
          if (decision !== 'Accept') {
            return null;
          }
        }

        if (internal === 'Addendum') {
          const decision = String(item?.RevisionIncludesChangeInTenderDo || '').trim().toLowerCase();
          console.log('🍩🤪decision ', decision, '\n 🫷🏻item?.RevisionIncludesChangeInTenderDo', item?.RevisionIncludesChangeInTenderDo);
          if ((tenderPhaseStr.indexOf('phase 2') > -1 && decision != 'yes') || tenderPhaseStr.indexOf('phase 1') > -1) {
            console.log("1", tenderPhaseStr.indexOf('phase 2') > -1 );
            console.log("2", decision != 'yes')
            console.log("3", tenderPhaseStr.indexOf('phase 1') > -1);
            console.log("🧻 null is returned");
            return null;
          }
        }

        if (internal === 'TenderCommitteeApprovalDate') {
          console.log("🤩 2");
          const decision = String(item?.RevisionIncludesChangeInTenderDo || '').trim();
          console.log('🍩🎈decision ', decision);
          if (decision != 'Yes') {
            console.log("🐴🐴🐴 decision != 'true' editefileds");
            return null;
          }

          if(displayRevisionIncludesChangeInTenderDo != true){
            console.log("🐴🐴🐴  displayRevisionIncludesChangeInTenderDo != true editefileds");
            return null;
          }
        }

        if (internal === 'StatusOfRFCresponseOrTcRFC') {
          const decision = String(item?.DecisionRegardingProposedChange || '').trim();
          
          if (decision !== 'Accept') {
            console.log('🍩🧋decision', decision);
            return null;
          }
         
        }

        if(internal === "StatusOfRFCresponseOrTcRFC" || internal === "TenderCommitteeApprovalDate" || internal === "Addendum"){
          console.log("🤩 2");
          if (tenderPhaseStr == 'phase 1 - bidders’ requests for clarifications (rfcs) of tender documents' || tenderPhaseStr.indexOf('phase 1') != -1) {
            console.log(
              "🍩🧋 PASE 1 StatusOfRFCresponseOrTcRFC, TenderCommitteeApprovalDate, Addendum "     
            );
            return null;
          }
        }

        if(internal === 'IntegrationTeamDecisionEditingLa'){
          return null;
        }

        if (isUrl(info)) {
          const raw = item?.[internal];

          const currentUrl =
            raw && typeof raw === 'object'
              ? (raw.Url || raw.url || '')
              : String(raw || '');

          const currentDesc =
            raw && typeof raw === 'object'
              ? (raw.Description || raw.desc || '')
              : '';
          console.log("placeholderMap?.[internal] ", placeholderMap?.[internal], "| placeholderMap ", placeholderMap);

          return (
            <div key={internal} style={{ marginBottom: 12 }}>
              <TextField
                label={label + ' – URL'}
                value={currentUrl}
                onGetErrorMessage={(value) => {
                  const v = (value || '').trim();
                  if (!v) return ''; // שדה ריק – לא מציגים שגיאה
                  const ok = /^https?:\/\/\S+$/i.test(v);
                  return ok ? '' : 'Enter a valid URL.';
                }}
                onChange={(_, v) => {
                  const urlStr = String(v || '').trim();
                  const descStr =  currentDesc || placeholderMap?.[internal] || urlStr;


                  onChange(
                    internal,
                    urlStr
                      ? { Url: urlStr, Description: descStr }
                      : null
                  );
                }}
                placeholder="https://example.com"
              />

              <TextField
                label={label + ' - Alternative text'}
                value={currentDesc}
                onChange={(_, v) => {
                  const descStr = String(v || '')
                  const urlStr =
                    raw && typeof raw === 'object'
                      ? (raw.Url || raw.url || '')
                      : currentUrl;

                  onChange(
                    internal,
                    urlStr
                      ? { Url: urlStr, Description: descStr || urlStr }
                      : null
                  );
                }}
                placeholder={placeholderMap?.[internal] || ""}
                styles={{ root: { marginTop: 4 } }}
              />

              {currentUrl && (
                <div style={{ marginTop: 4 }}>
                  <a href={currentUrl} target="_blank" rel="noreferrer">
                    {currentDesc || currentUrl}
                  </a>
                </div>
              )}
            </div>
          );
        }


        // Boolean
        if (isBoolean(info)) {
          return (
            <Toggle
              key={internal}
              label={label}
              checked={!!val}
              onChange={editable ? ((_, c) => onChange(internal, !!c)) : undefined}
              inlineLabel
              disabled={!editable}
            />
          );
        }

        // Date/DateTime
        if (isDate(info)) {
          return (
            <DatePicker
              key={internal}
              label={label}
              value={val ? new Date(val) : undefined}
              onSelectDate={editable ? ((d) => onChange(internal, d ? d.toISOString() : null)) : undefined}
              allowTextInput
              disabled={!editable}
              isRequired={required}
              styles={datePickerStyles as any}
            />
          );
        }

        // Choice / MultiChoice
        if (isChoice(info)) {
          const isMulti = isMultiField(info);

          const rawChoices: string[] = (function () {
            if (!info) return [];
            if (Array.isArray(info.Choices)) return info.Choices as string[];
            if (info.Choices && Array.isArray(info.Choices.results)) return info.Choices.results as string[];
            return [];
          })();

          const current = val;
          const isIntegrationDecision = internal === 'IntegrationTeamDecisionImplement';
          const appliesVal = String(item?.['DecisionAppliesToOtherWorksTende'] ?? '').trim();
          const allowByDependency = (appliesVal === 'All Infra 1 tenders' || appliesVal === 'Infra#1 DB - M3-WP2'|| appliesVal === 'Infra#1 DB - M2-WP3'
            || appliesVal === 'Infra#1 DB - M1-WP1 + WP2' || appliesVal ==='M3-WPO (Outer Boxes)' ||
            appliesVal.indexOf('All Infra 1 tenders') !== -1 ||
            appliesVal.indexOf('Infra#1 DB - M3-WP2') !== -1 || appliesVal.indexOf('Infra#1 DB - M2-WP3') !== -1 ||
              appliesVal.indexOf('Infra#1 DB - M1-WP1 + WP2') !== -1 || appliesVal.indexOf('M3-WPO (Outer Boxes)') !== -1);
          const decisionDate = formatDateDDMMYYYY(
            item?.['DateForIntegrationTeamDecisionIm']
              ? String(item['DateForIntegrationTeamDecisionIm']).trim()
              : ''
          );

          const disabledComputed = isIntegrationDecision
            ? (!allowByDependency || !editable)
            : (!editable);

          /*
          const options: IDropdownOption[] = (choiceOverrides[internal] || rawChoices).map((ch: any) => {
            if (typeof ch === 'string') {
              if (isIntegrationDecision && (ch === 'Done' || ch === 'Pending') && decisionDate) {
                return { key: ch, text: `${ch} (${decisionDate})` };
              }
              return { key: ch, text: ch };
            }
            return ch;
          });*/
          const baseChoices: any[] =
            (choiceOverrides && choiceOverrides[internal])
              ? choiceOverrides[internal]
              : rawChoices;

          const options: IDropdownOption[] = baseChoices.map((ch: any) => {
            if (typeof ch === 'string') {
              if (isIntegrationDecision && (ch === 'Done' || ch === 'Pending') && decisionDate) {
                return { key: ch, text: `${ch} (${decisionDate})` };
              }
              return { key: ch, text: ch };
            }
            return ch; // אם כבר הגיע כ-IDropdownOption
          });

          const selectedKeys = isMulti ? (Array.isArray(current) ? current : []) : undefined;
          const selectedKey = !isMulti ? (current === null || current === undefined ? undefined : String(current)) : undefined;

          return (
            <Dropdown
              key={internal}
              label={label}
              placeholder={isMulti ? 'Select values' : 'Select value'}
              options={options}
              multiSelect={isMulti}
              selectedKeys={isMulti ? selectedKeys : undefined}
              selectedKey={!isMulti ? selectedKey : undefined}
              disabled={disabledComputed}
              required={required}
              errorMessage={error}
              styles={dropdownStyles as any}
              onChange={(!disabledComputed) ? ((_, opt) => {
                if (!opt) return;
                if (isMulti) {
                  const prev = Array.isArray(current) ? [...current] : [];
                  const k = String(opt.key);
                  let idx = -1;
                  for (let i = 0; i < prev.length; i++) {
                    if (String(prev[i]) === k) { idx = i; break; }
                  }
                  if (opt.selected && idx < 0) prev.push(k);
                  if (!opt.selected && idx > -1) prev.splice(idx, 1);
                  onChange(internal, prev);
                } else {
                  onChange(internal, String(opt.key));
                }
              }) : undefined}
            />
          );
        }

        // Lookup/User
        if (isLookupOrUser(info)) {
          const isMulti = isMultiField(info);
          return (
            <TextField
              key={internal}
              label={label}
              description={hint || (isMulti ? 'Enter identifiers separated by commas, for example: 12,34,56' : 'Enter my numeric ID')}
              value={
                isMulti
                  ? (Array.isArray(val) ? (val as any[]).join(',') : '')
                  : (val === null || val === undefined ? '' : String(val))
              }
              disabled={!editable}
              onGetErrorMessage={editable ? ((t) => {
                if (!t) return '';
                if (isMulti) {
                  const parts = String(t).split(',');
                  for (let i = 0; i < parts.length; i++) {
                    const x: string = parts[i].trim();
                    if (!x) continue;
                    if (isNaN(Number(x))) return 'Enter only numbers separated by commas';
                  }
                  return '';
                } else {
                  return isNaN(Number(t as any)) ? 'Please enter a valid number' : '';
                }
              }) : undefined}
              required={required}
              errorMessage={error}
              onChange={editable ? ((_, t) => {
                if (isMulti) {
                  const nums: number[] = [];
                  const parts = String(t || '').split(',');
                  for (let i = 0; i < parts.length; i++) {
                    const x = parts[i].trim();
                    if (!x) continue;
                    const n = Number(x);
                    if (!isNaN(n)) nums.push(n);
                  }
                  onChange(internal, nums);
                } else {
                  const trimmed = (t || '').trim();
                  if (trimmed === '') onChange(internal, null);
                  else {
                    const n = Number(trimmed);
                    onChange(internal, isNaN(n) ? null : n);
                  }
                }
              }) : undefined}
              styles={{ ...textStyles, root: { maxWidth: baseControlMaxWidth } } as any}
            />
          );
        }

        // Number/Currency
        if (isNumber(info)) {
          return (
            <TextField
              key={internal}
              label={label}
              description={hint}
              type="number"
              value={val === null || val === undefined ? '' : String(val)}
              disabled={!editable}
              onGetErrorMessage={editable ? ((t) => {
                if (!t) return '';
                const n = toNumberOrNull(t);
                return n === null ? 'Please enter a valid number' : '';
              }) : undefined}
              required={required}
              errorMessage={error}
              onChange={editable ? ((_, t) => onChange(internal, toNumberOrNull(t))) : undefined}
              styles={{ ...textStyles, root: { maxWidth: baseControlMaxWidth } } as any}
            />
          );
        }

        // Rich Text → טקסט רגיל/עריך
        if (isRichText(info)) {
          if (!editable) {
            
            return (
              <div
                key={internal}
                style={{
                  padding: 10,
                  background: 'rgba(15,23,42,0.02)',
                  borderRadius: 12,
                  border: '1px solid #e5e7eb'
                }}
              >
                <div style={{ fontSize: 12, color: '#6b7280', marginBottom: 4 }}>{label}</div>
                <div style={{ whiteSpace: 'pre-wrap', fontWeight: 600 }}>
                  {htmlToPlainText(val)}
                </div>
              </div>
            );
          }
          const plain = htmlToPlainText(val);

          return (
            <TextField
              key={internal}
              label={label}
              description={hint}
              multiline
              autoAdjustHeight
              value={htmlToPlainText(val)}
               placeholder={
                placeholder
                  ? placeholder
                  : (plain.trim() === '' ? (placeholderMap?.[internal] ?? undefined) : undefined)
              }
              /*
              placeholder=
                (!val || String(htmlToPlainText(val)).trim() === '')
                  ? (placeholderMap?.[internal] || undefined)
                  : undefined
              }*/
              required={required}
              errorMessage={error}
              onChange={(_, t) => onChange(internal, t ?? '')}
              styles={{ ...textStyles, root: { maxWidth: baseControlMaxWidth } } as any}
            />
          );
        }
       
        // Taxonomy (Placeholder)
        if (isTaxonomy(info)) {
          return (
            <TextField
              key={internal}
              label={label}
              description={hint || 'Taxonomy field – a dedicated component later'}
              readOnly
              value={val ? String(val) : ''}
              styles={{ ...textStyles, root: { maxWidth: baseControlMaxWidth } } as any}
            />
          );
        }

       
        // Text רגיל
        console.log("🐢 this is finalt a textfield internal ", internal, "place holder ", placeholderMap?.[internal] );
        return (
          <TextField
            key={internal}
            label={label}
            description={hint}
            value={val === null || val === undefined ? '' : String(val)}
            readOnly={!editable}
            placeholder={placeholderMap?.[internal]}
            required={required}
            errorMessage={error}
            onChange={editable ? ((_, nv) => onChange(internal, nv)) : undefined}
            styles={{ ...textStyles, root: { maxWidth: baseControlMaxWidth } } as any}
          />
        );

      })}
    </Stack>
  );
};

export default EditableFields;
