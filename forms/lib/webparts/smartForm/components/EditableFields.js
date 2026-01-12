var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __spreadArray = (this && this.__spreadArray) || function (to, from, pack) {
    if (pack || arguments.length === 2) for (var i = 0, l = from.length, ar; i < l; i++) {
        if (ar || !(i in from)) {
            if (!ar) ar = Array.prototype.slice.call(from, 0, i);
            ar[i] = from[i];
        }
    }
    return to.concat(ar || Array.prototype.slice.call(from));
};
// src/webparts/smartForm/components/EditableFields.tsx
import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown } from '@fluentui/react/lib/Dropdown';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { DatePicker } from '@fluentui/react/lib/DatePicker';
import { isSystemField } from '../shared/constants';
var REQUIRED_FIELDS = [
    'DecisionRegardingProposedChange',
    'DecisionAppliesToOtherWorksTende',
    'RevisionIncludesChangeInTenderDo',
    'DecisionDate',
    'DueDateCalculated',
    'ActualDate'
];
// ===== עזרי טקסט =====
var lpad = function (s, len, ch) {
    if (ch === void 0) { ch = '0'; }
    var t = String(s == null ? '' : s);
    while (t.length < len)
        t = ch + t;
    return t;
};
export var formatDateDDMMYYYY = function (d) {
    if (!d)
        return '';
    var dt = typeof d === 'string' ? new Date(d) : d;
    if (!(dt instanceof Date) || isNaN(dt.getTime()))
        return '';
    var dd = lpad(dt.getDate(), 2);
    var mm = lpad(dt.getMonth() + 1, 2);
    var yyyy = dt.getFullYear();
    return "".concat(dd, "/").concat(mm, "/").concat(yyyy);
};
var strHas = function (a, frag) {
    return String(a || '').toLowerCase().indexOf(String(frag || '').toLowerCase()) > -1;
};
function htmlToPlainText(html) {
    if (!html)
        return '';
    var s = String(html);
    s = s
        .replace(/<\s*li[^>]*>/gi, '• ')
        .replace(/<\s*br\s*\/?>/gi, '\n')
        .replace(/<\/\s*p\s*>/gi, '\n')
        .replace(/<\/\s*li\s*>/gi, '\n');
    s = s.replace(/<[^>]+>/g, '');
    var ta = document.createElement('textarea');
    ta.innerHTML = s;
    return ta.value;
}
// ===== זיהוי סוגי שדות =====
function isMultiField(info) {
    if (!info)
        return false;
    if (info.AllowMultipleValues === true)
        return true;
    return strHas(info === null || info === void 0 ? void 0 : info.TypeAsString, 'multi');
}
function isLookupOrUser(info) {
    if (!info || !info.TypeAsString)
        return false;
    var t = String(info.TypeAsString || '').toLowerCase();
    return t.indexOf('lookup') > -1 || t.indexOf('user') > -1;
}
function isChoice(info) {
    if (!info || !info.TypeAsString)
        return false;
    var t = String(info.TypeAsString || '').toLowerCase();
    return t === 'choice' || t === 'multichoice';
}
function isTaxonomy(info) {
    if (!info || !info.TypeAsString)
        return false;
    return strHas(info.TypeAsString, 'taxonomy');
}
function isDate(info) {
    if (!info || !info.TypeAsString)
        return false;
    var t = String(info.TypeAsString || '').toLowerCase();
    return t.indexOf('date') > -1;
}
function isNumber(info) {
    if (!info || !info.TypeAsString)
        return false;
    var t = String(info.TypeAsString || '').toLowerCase();
    return t === 'number' || t === 'currency';
}
function isBoolean(info) {
    if (!info || !info.TypeAsString)
        return false;
    return String(info.TypeAsString || '').toLowerCase() === 'boolean';
}
function isRichText(info) {
    if (!info)
        return false;
    return info.RichText === true || String(info.TypeAsString || '').toLowerCase() === 'note';
}
function isUrl(info) {
    if (!info || !info.TypeAsString)
        return false;
    var t = String(info.TypeAsString || '').toLowerCase();
    return t === 'url';
}
// ===== טקסט עזרה =====
function helpTextFor(info) {
    if (!info)
        return undefined;
    if (isBoolean(info))
        return 'Binary value: yes/no.';
    if (isDate(info))
        return 'Date/Date-Time (can be selected from the date range).';
    if (isNumber(info))
        return 'Number (digits and decimal point).';
    if (isChoice(info) && isMultiField(info))
        return 'Multiple selection from a list of options.';
    if (isChoice(info))
        return 'A single selection from a list of options.';
    if (isLookupOrUser(info) && isMultiField(info))
        return 'Link/User – Identifiers (Id) separated by comma.';
    if (isLookupOrUser(info))
        return 'Link/User – Unique Identifier (Id).';
    if (isRichText(info))
        return 'Rich text will be displayed as plain text (without HTML tags).';
    if (isTaxonomy(info))
        return 'Taxonomy field.';
    return undefined;
}
// ===== עזר מספרי =====
function toNumberOrNull(s) {
    if (s === undefined || s === null)
        return null;
    var t = String(s).trim();
    if (!t)
        return null;
    var normalized = t.replace(',', '.');
    var n = parseFloat(normalized);
    return isFinite(n) ? n : null;
}
// ===== עיצוב בסיסי לכל הקונטרולים =====
var baseControlMaxWidth = 480;
var textStyles = {
    label: { fontWeight: 600 },
    fieldGroup: { borderRadius: 10 },
};
var dropdownStyles = {
    label: { fontWeight: 600 },
    root: { maxWidth: baseControlMaxWidth },
};
var datePickerStyles = {
    root: { maxWidth: baseControlMaxWidth },
    textField: __assign({}, textStyles),
};
// ===== הקומפוננטה =====
var EditableFields = function (_a) {
    var item = _a.item, onChange = _a.onChange, _b = _a.fieldOrder, fieldOrder = _b === void 0 ? [] : _b, _c = _a.hideFields, hideFields = _c === void 0 ? [] : _c, _d = _a.internalToTitle, internalToTitle = _d === void 0 ? {} : _d, _e = _a.fieldInfoMap, fieldInfoMap = _e === void 0 ? {} : _e, _f = _a.canEdit, canEdit = _f === void 0 ? function () { return false; } : _f, _g = _a.placeholderMap, placeholderMap = _g === void 0 ? {} : _g, _h = _a.choiceOverrides, choiceOverrides = _h === void 0 ? {} : _h, tenderPhase = _a.tenderPhase, _j = _a.requiredMap, requiredMap = _j === void 0 ? {} : _j, _k = _a.errorMap, errorMap = _k === void 0 ? {} : _k, _l = _a.labelOverrides, labelOverrides = _l === void 0 ? {} : _l;
    var extraHide = ['Integration', 'IntegrationId', 'Id', 'ID', 'Title'];
    console.log("🥫 fieldOrder ", fieldOrder, "\n🍼 hideFields", hideFields);
    var keys = React.useMemo(function () {
        console.log("🎋Object.keys(item)", Object.keys(item));
        var all = Object.keys(item || {}).filter(function (k) {
            return hideFields.indexOf(k) === -1 &&
                extraHide.indexOf(k) === -1 &&
                !isSystemField(k) &&
                fieldOrder.indexOf(k) > -1;
        });
        console.log("all- 1 🔑 ", all);
        var ordered = [];
        var seen = {};
        for (var i = 0; i < fieldOrder.length; i++) {
            var f = fieldOrder[i];
            if (all.indexOf(f) > -1 && !seen[f]) {
                ordered.push(f);
                seen[f] = true;
            }
        }
        for (var i = 0; i < all.length; i++) {
            var f = all[i];
            if (!seen[f]) {
                ordered.push(f);
                seen[f] = true;
            }
        }
        console.log("all 🔑 ", all);
        console.log("ordered 🔑 ", ordered);
        return ordered;
    }, [item, fieldOrder, hideFields]);
    var displayRevisionIncludesChangeInTenderDo = true;
    return (React.createElement(Stack, { tokens: { childrenGap: 12 }, styles: {
            root: {
                paddingTop: 4,
            }
        } }, keys.map(function (internal) {
        var _a, _b;
        if (!internal)
            return null;
        var tenderPhaseStr = String(tenderPhase || '').trim().toLowerCase();
        console.log("🍩🧋 tenderPhaseStr ", tenderPhaseStr);
        var val = item[internal];
        var info = fieldInfoMap[internal];
        console.log("🫏 const label = (labelOverrides && labelOverrides[internal]) || internalToTitle[internal] ||internal; ", "label  ", (labelOverrides && labelOverrides[internal]) || internalToTitle[internal] || internal, "labelOverrides  ", labelOverrides, "  labelOverrides[internal] ", labelOverrides[internal], "internalToTitle[internal]", internalToTitle[internal], "  internal   ", internal);
        var label = (labelOverrides && labelOverrides[internal]) || (internalToTitle === null || internalToTitle === void 0 ? void 0 : internalToTitle[internal]);
        //|| internalToTitle[internal] || internal;
        var hint = helpTextFor(info);
        var editable = canEdit ? !!canEdit(internal) : true;
        var required = editable && REQUIRED_FIELDS.indexOf(internal) > -1;
        var error = errorMap[internal];
        console.log('🤗tenderPhaseLc ', tenderPhaseStr);
        var placeholder = placeholderMap === null || placeholderMap === void 0 ? void 0 : placeholderMap[internal];
        if (internal === 'RFCorTcRFCasPublishedByNTaToBeFi') {
            var decision = String((item === null || item === void 0 ? void 0 : item.DecisionRegardingProposedChange) || '').trim();
            console.log('🍩decision ', decision);
            if (tenderPhaseStr.indexOf('phase 1') > -1) {
                return null;
            }
            if (decision === 'Accept') {
                placeholder = 'Enter final version for publication here';
            }
        }
        if (internal === 'Assignedto' || internal === 'SubCategory') {
            if (tenderPhaseStr.indexOf('phase 1') === -1) {
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
        if (internal === 'RevisionIncludesChangeInTenderDo' ||
            internal === 'RFCResponseLetterNo' ||
            internal === 'RFCresponseAsPublishedToBeFilled') {
            if (tenderPhaseStr != 'phase 2 - bidders’ requests for clarifications (rfcs) of tender documents' && tenderPhaseStr.indexOf('phase 2') === -1) {
                if (internal === 'RevisionIncludesChangeInTenderDo') {
                    displayRevisionIncludesChangeInTenderDo = false;
                }
                return null;
            }
        }
        if (internal === 'RFCresponseAsPublishedToBeFilled') {
            var decision = String((item === null || item === void 0 ? void 0 : item.DecisionRegardingProposedChange) || '').trim();
            console.log('🍩decision ', decision);
            if (decision !== 'Accept') {
                return null;
            }
        }
        if (internal === 'Addendum') {
            var decision = String((item === null || item === void 0 ? void 0 : item.RevisionIncludesChangeInTenderDo) || '').trim().toLowerCase();
            console.log('🍩🤪decision ', decision, '\n 🫷🏻item?.RevisionIncludesChangeInTenderDo', item === null || item === void 0 ? void 0 : item.RevisionIncludesChangeInTenderDo);
            if ((tenderPhaseStr.indexOf('phase 2') > -1 && decision != 'yes') || tenderPhaseStr.indexOf('phase 1') > -1) {
                console.log("1", tenderPhaseStr.indexOf('phase 2') > -1);
                console.log("2", decision != 'yes');
                console.log("3", tenderPhaseStr.indexOf('phase 1') > -1);
                console.log("🧻 null is returned");
                return null;
            }
        }
        if (internal === 'TenderCommitteeApprovalDate') {
            console.log("🤩 2");
            var decision = String((item === null || item === void 0 ? void 0 : item.RevisionIncludesChangeInTenderDo) || '').trim();
            console.log('🍩🎈decision ', decision);
            if (decision != 'Yes') {
                console.log("🐴🐴🐴 decision != 'true' editefileds");
                return null;
            }
            if (displayRevisionIncludesChangeInTenderDo != true) {
                console.log("🐴🐴🐴  displayRevisionIncludesChangeInTenderDo != true editefileds");
                return null;
            }
        }
        if (internal === 'StatusOfRFCresponseOrTcRFC') {
            var decision = String((item === null || item === void 0 ? void 0 : item.DecisionRegardingProposedChange) || '').trim();
            if (decision !== 'Accept') {
                console.log('🍩🧋decision', decision);
                return null;
            }
        }
        if (internal === "StatusOfRFCresponseOrTcRFC" || internal === "TenderCommitteeApprovalDate" || internal === "Addendum") {
            console.log("🤩 2");
            if (tenderPhaseStr == 'phase 1 - bidders’ requests for clarifications (rfcs) of tender documents' || tenderPhaseStr.indexOf('phase 1') != -1) {
                console.log("🍩🧋 PASE 1 StatusOfRFCresponseOrTcRFC, TenderCommitteeApprovalDate, Addendum ");
                return null;
            }
        }
        if (internal === 'IntegrationTeamDecisionEditingLa') {
            return null;
        }
        if (isUrl(info)) {
            var raw_1 = item === null || item === void 0 ? void 0 : item[internal];
            var currentUrl_1 = raw_1 && typeof raw_1 === 'object'
                ? (raw_1.Url || raw_1.url || '')
                : String(raw_1 || '');
            var currentDesc_1 = raw_1 && typeof raw_1 === 'object'
                ? (raw_1.Description || raw_1.desc || '')
                : '';
            console.log("placeholderMap?.[internal] ", placeholderMap === null || placeholderMap === void 0 ? void 0 : placeholderMap[internal], "| placeholderMap ", placeholderMap);
            return (React.createElement("div", { key: internal, style: { marginBottom: 12 } },
                React.createElement(TextField, { label: label + ' – URL', value: currentUrl_1, onGetErrorMessage: function (value) {
                        var v = (value || '').trim();
                        if (!v)
                            return ''; // שדה ריק – לא מציגים שגיאה
                        var ok = /^https?:\/\/\S+$/i.test(v);
                        return ok ? '' : 'Enter a valid URL.';
                    }, onChange: function (_, v) {
                        var urlStr = String(v || '').trim();
                        var descStr = currentDesc_1 || (placeholderMap === null || placeholderMap === void 0 ? void 0 : placeholderMap[internal]) || urlStr;
                        onChange(internal, urlStr
                            ? { Url: urlStr, Description: descStr }
                            : null);
                    }, placeholder: "https://example.com" }),
                React.createElement(TextField, { label: label + ' - Alternative text', value: currentDesc_1, onChange: function (_, v) {
                        var descStr = String(v || '');
                        var urlStr = raw_1 && typeof raw_1 === 'object'
                            ? (raw_1.Url || raw_1.url || '')
                            : currentUrl_1;
                        onChange(internal, urlStr
                            ? { Url: urlStr, Description: descStr || urlStr }
                            : null);
                    }, placeholder: (placeholderMap === null || placeholderMap === void 0 ? void 0 : placeholderMap[internal]) || "", styles: { root: { marginTop: 4 } } }),
                currentUrl_1 && (React.createElement("div", { style: { marginTop: 4 } },
                    React.createElement("a", { href: currentUrl_1, target: "_blank", rel: "noreferrer" }, currentDesc_1 || currentUrl_1)))));
        }
        // Boolean
        if (isBoolean(info)) {
            return (React.createElement(Toggle, { key: internal, label: label, checked: !!val, onChange: editable ? (function (_, c) { return onChange(internal, !!c); }) : undefined, inlineLabel: true, disabled: !editable }));
        }
        // Date/DateTime
        if (isDate(info)) {
            return (React.createElement(DatePicker, { key: internal, label: label, value: val ? new Date(val) : undefined, onSelectDate: editable ? (function (d) { return onChange(internal, d ? d.toISOString() : null); }) : undefined, allowTextInput: true, disabled: !editable, isRequired: required, styles: datePickerStyles }));
        }
        // Choice / MultiChoice
        if (isChoice(info)) {
            var isMulti_1 = isMultiField(info);
            var rawChoices = (function () {
                if (!info)
                    return [];
                if (Array.isArray(info.Choices))
                    return info.Choices;
                if (info.Choices && Array.isArray(info.Choices.results))
                    return info.Choices.results;
                return [];
            })();
            var current_1 = val;
            var isIntegrationDecision_1 = internal === 'IntegrationTeamDecisionImplement';
            var appliesVal = String((_a = item === null || item === void 0 ? void 0 : item['DecisionAppliesToOtherWorksTende']) !== null && _a !== void 0 ? _a : '').trim();
            var allowByDependency = (appliesVal === 'All Infra 1 tenders' || appliesVal === 'Infra#1 DB - M3-WP2' || appliesVal === 'Infra#1 DB - M2-WP3'
                || appliesVal === 'Infra#1 DB - M1-WP1 + WP2' || appliesVal === 'M3-WPO (Outer Boxes)' ||
                appliesVal.indexOf('All Infra 1 tenders') !== -1 ||
                appliesVal.indexOf('Infra#1 DB - M3-WP2') !== -1 || appliesVal.indexOf('Infra#1 DB - M2-WP3') !== -1 ||
                appliesVal.indexOf('Infra#1 DB - M1-WP1 + WP2') !== -1 || appliesVal.indexOf('M3-WPO (Outer Boxes)') !== -1);
            var decisionDate_1 = formatDateDDMMYYYY((item === null || item === void 0 ? void 0 : item['DateForIntegrationTeamDecisionIm'])
                ? String(item['DateForIntegrationTeamDecisionIm']).trim()
                : '');
            var disabledComputed = isIntegrationDecision_1
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
            var baseChoices = (choiceOverrides && choiceOverrides[internal])
                ? choiceOverrides[internal]
                : rawChoices;
            var options = baseChoices.map(function (ch) {
                if (typeof ch === 'string') {
                    if (isIntegrationDecision_1 && (ch === 'Done' || ch === 'Pending') && decisionDate_1) {
                        return { key: ch, text: "".concat(ch, " (").concat(decisionDate_1, ")") };
                    }
                    return { key: ch, text: ch };
                }
                return ch; // אם כבר הגיע כ-IDropdownOption
            });
            var selectedKeys = isMulti_1 ? (Array.isArray(current_1) ? current_1 : []) : undefined;
            var selectedKey = !isMulti_1 ? (current_1 === null || current_1 === undefined ? undefined : String(current_1)) : undefined;
            return (React.createElement(Dropdown, { key: internal, label: label, placeholder: isMulti_1 ? 'Select values' : 'Select value', options: options, multiSelect: isMulti_1, selectedKeys: isMulti_1 ? selectedKeys : undefined, selectedKey: !isMulti_1 ? selectedKey : undefined, disabled: disabledComputed, required: required, errorMessage: error, styles: dropdownStyles, onChange: (!disabledComputed) ? (function (_, opt) {
                    if (!opt)
                        return;
                    if (isMulti_1) {
                        var prev = Array.isArray(current_1) ? __spreadArray([], current_1, true) : [];
                        var k = String(opt.key);
                        var idx = -1;
                        for (var i = 0; i < prev.length; i++) {
                            if (String(prev[i]) === k) {
                                idx = i;
                                break;
                            }
                        }
                        if (opt.selected && idx < 0)
                            prev.push(k);
                        if (!opt.selected && idx > -1)
                            prev.splice(idx, 1);
                        onChange(internal, prev);
                    }
                    else {
                        onChange(internal, String(opt.key));
                    }
                }) : undefined }));
        }
        // Lookup/User
        if (isLookupOrUser(info)) {
            var isMulti_2 = isMultiField(info);
            return (React.createElement(TextField, { key: internal, label: label, description: hint || (isMulti_2 ? 'Enter identifiers separated by commas, for example: 12,34,56' : 'Enter my numeric ID'), value: isMulti_2
                    ? (Array.isArray(val) ? val.join(',') : '')
                    : (val === null || val === undefined ? '' : String(val)), disabled: !editable, onGetErrorMessage: editable ? (function (t) {
                    if (!t)
                        return '';
                    if (isMulti_2) {
                        var parts = String(t).split(',');
                        for (var i = 0; i < parts.length; i++) {
                            var x = parts[i].trim();
                            if (!x)
                                continue;
                            if (isNaN(Number(x)))
                                return 'Enter only numbers separated by commas';
                        }
                        return '';
                    }
                    else {
                        return isNaN(Number(t)) ? 'Please enter a valid number' : '';
                    }
                }) : undefined, required: required, errorMessage: error, onChange: editable ? (function (_, t) {
                    if (isMulti_2) {
                        var nums = [];
                        var parts = String(t || '').split(',');
                        for (var i = 0; i < parts.length; i++) {
                            var x = parts[i].trim();
                            if (!x)
                                continue;
                            var n = Number(x);
                            if (!isNaN(n))
                                nums.push(n);
                        }
                        onChange(internal, nums);
                    }
                    else {
                        var trimmed = (t || '').trim();
                        if (trimmed === '')
                            onChange(internal, null);
                        else {
                            var n = Number(trimmed);
                            onChange(internal, isNaN(n) ? null : n);
                        }
                    }
                }) : undefined, styles: __assign(__assign({}, textStyles), { root: { maxWidth: baseControlMaxWidth } }) }));
        }
        // Number/Currency
        if (isNumber(info)) {
            return (React.createElement(TextField, { key: internal, label: label, description: hint, type: "number", value: val === null || val === undefined ? '' : String(val), disabled: !editable, onGetErrorMessage: editable ? (function (t) {
                    if (!t)
                        return '';
                    var n = toNumberOrNull(t);
                    return n === null ? 'Please enter a valid number' : '';
                }) : undefined, required: required, errorMessage: error, onChange: editable ? (function (_, t) { return onChange(internal, toNumberOrNull(t)); }) : undefined, styles: __assign(__assign({}, textStyles), { root: { maxWidth: baseControlMaxWidth } }) }));
        }
        // Rich Text → טקסט רגיל/עריך
        if (isRichText(info)) {
            if (!editable) {
                return (React.createElement("div", { key: internal, style: {
                        padding: 10,
                        background: 'rgba(15,23,42,0.02)',
                        borderRadius: 12,
                        border: '1px solid #e5e7eb'
                    } },
                    React.createElement("div", { style: { fontSize: 12, color: '#6b7280', marginBottom: 4 } }, label),
                    React.createElement("div", { style: { whiteSpace: 'pre-wrap', fontWeight: 600 } }, htmlToPlainText(val))));
            }
            var plain = htmlToPlainText(val);
            return (React.createElement(TextField, { key: internal, label: label, description: hint, multiline: true, autoAdjustHeight: true, value: htmlToPlainText(val), placeholder: placeholder
                    ? placeholder
                    : (plain.trim() === '' ? ((_b = placeholderMap === null || placeholderMap === void 0 ? void 0 : placeholderMap[internal]) !== null && _b !== void 0 ? _b : undefined) : undefined), 
                /*
                placeholder=
                  (!val || String(htmlToPlainText(val)).trim() === '')
                    ? (placeholderMap?.[internal] || undefined)
                    : undefined
                }*/
                required: required, errorMessage: error, onChange: function (_, t) { return onChange(internal, t !== null && t !== void 0 ? t : ''); }, styles: __assign(__assign({}, textStyles), { root: { maxWidth: baseControlMaxWidth } }) }));
        }
        // Taxonomy (Placeholder)
        if (isTaxonomy(info)) {
            return (React.createElement(TextField, { key: internal, label: label, description: hint || 'Taxonomy field – a dedicated component later', readOnly: true, value: val ? String(val) : '', styles: __assign(__assign({}, textStyles), { root: { maxWidth: baseControlMaxWidth } }) }));
        }
        // Text רגיל
        console.log("🐢 this is finalt a textfield internal ", internal, "place holder ", placeholderMap === null || placeholderMap === void 0 ? void 0 : placeholderMap[internal]);
        return (React.createElement(TextField, { key: internal, label: label, description: hint, value: val === null || val === undefined ? '' : String(val), readOnly: !editable, placeholder: placeholderMap === null || placeholderMap === void 0 ? void 0 : placeholderMap[internal], required: required, errorMessage: error, onChange: editable ? (function (_, nv) { return onChange(internal, nv); }) : undefined, styles: __assign(__assign({}, textStyles), { root: { maxWidth: baseControlMaxWidth } }) }));
    })));
};
export default EditableFields;
//# sourceMappingURL=EditableFields.js.map