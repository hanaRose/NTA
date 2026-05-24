//src\webparts\smartForm\components\FormApp.tsx
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
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import { Stack, Pivot, PivotItem, PrimaryButton, DefaultButton, MessageBar, MessageBarType, ComboBox, Dropdown } from '@fluentui/react';
import { getSP } from './pnpjsConfig';
import { isSystemField } from '../shared/constants';
import { INTEGRATION_LIST_ID, INTEGRATION_PREVIEW_FIELD_INTERNAL } from '../shared/internalNames';
import { 
//getFieldMapsByTitle,
//getFieldInfoMap,
getFieldInfoMapById, fetchIntegrationItemByGuid, fetchOrCreatePmoByIntegration, savePmoItem, loadFieldPermissionMap, loadGeneralRoleUsers, loadTenderTeamUsersFromIntegration, canUserEditField, getFieldMapsById } from '../shared/data';
import stepsConfigJson from '../stepsConfig.json';
import '@pnp/sp/site-users/web';
import EditableFields from './EditableFields';
import '@pnp/sp/views';
import '@pnp/sp/lists';
// ===== עיצוב בסיסי (צבעים, כרטיסים וכו') =====
var PAGE_BG = 'linear-gradient(135deg, #f4f6fb 0%, #e7f2ff 40%, #f9fafb 100%)';
var CARD_BG = '#ffffff';
var CARD_SHADOW = '0 14px 40px rgba(15, 23, 42, 0.12)';
var CARD_RADIUS = 18;
var ACCENT = '#00498f';
//const ACCENT_SOFT = '#e6f0ff';// ===== Reload Guard (per-user) =====
var RELOAD_GUARD_LIST_TITLE = "ReloadGuard";
var RELOAD_GUARD_USER_FIELD = "User"; // Person field
var RELOAD_GUARD_FLAG_FIELD = "HasReloadedOnce"; // Boolean field
/**
 * 1) קוראת את כל המכרזים (Title) מרשימת SharePoint ומחזירה Set של שמות (trim).
 *    - עובדת גם אם יש יותר מ-5000 פריטים (ע"י paging).
 */
/*
export async function fetchAllTenderTitles(params: {
  sp: any;
  workTendersListId: string;     // GUID של הרשימה
  titleFieldInternalName?: string; // default: "LinkTitle"
}): Promise<Set<string>> {
  const { sp, workTendersListId, titleFieldInternalName = "LinkTitle" } = params;
  console.log("❤️‍🩹0");
  const list = sp.web.lists.getById(workTendersListId);
  console.log("❤️‍🩹1");
  const titles = new Set<string>();
  console.log("❤️‍🩹2");
  // paging (PnPjs)
  //const pageSize = 2000;
  console.log("❤️‍🩹3");
  let items = await list.items.select(titleFieldInternalName).getAll();;
  console.log("❤️‍🩹4");
  for (const item of items) {
    const t = String(item?.[titleFieldInternalName] ?? "").trim();
    if (t) titles.add(t);
  }

  return titles;
}
*/
export function fetchAllTenderTitles(params) {
    var _a;
    return __awaiter(this, void 0, void 0, function () {
        var sp, workTendersListId, _b, titleFieldInternalName, list, titles, pageSize, lastId, batch, _i, batch_1, item, t;
        return __generator(this, function (_c) {
            switch (_c.label) {
                case 0:
                    sp = params.sp, workTendersListId = params.workTendersListId, _b = params.titleFieldInternalName, titleFieldInternalName = _b === void 0 ? "LinkTitle" : _b;
                    list = sp.web.lists.getById(workTendersListId);
                    titles = new Set();
                    pageSize = 2000;
                    lastId = 0;
                    _c.label = 1;
                case 1:
                    if (!true) return [3 /*break*/, 3];
                    return [4 /*yield*/, list.items
                            .select("Id", titleFieldInternalName)
                            .filter("Id gt ".concat(lastId))
                            .orderBy("Id", true)
                            .top(pageSize)()];
                case 2:
                    batch = _c.sent();
                    for (_i = 0, batch_1 = batch; _i < batch_1.length; _i++) {
                        item = batch_1[_i];
                        t = String((_a = item === null || item === void 0 ? void 0 : item[titleFieldInternalName]) !== null && _a !== void 0 ? _a : "").trim();
                        if (t)
                            titles.add(t);
                    }
                    if (batch.length < pageSize)
                        return [3 /*break*/, 3];
                    lastId = batch[batch.length - 1].Id;
                    return [3 /*break*/, 1];
                case 3:
                    titles.add('');
                    return [2 /*return*/, titles];
            }
        });
    });
}
/**
 * 2) בודקת שהמחרוזת היא:
 *    - בדיוק "All Infra 1 tenders"
 *    - או בדיוק "Not relevant to additional tenders"
 *    - או רשימת ערכים מופרדת בפסיקים, כשכל ערך קיים ברשימת המכרזים (Title)
 *    - בלי "עוד מילים" ובלי ערכים לא מוכרים.
 */
export function isValidTenderSelection(params) {
    var input = params.input, validTitles = params.validTitles, _a = params.allowAllInfra, allowAllInfra = _a === void 0 ? true : _a, _b = params.allowNotRelevant, allowNotRelevant = _b === void 0 ? true : _b;
    var raw = String(input !== null && input !== void 0 ? input : "").trim();
    console.log("💒🛹🧸raw |", raw, "|");
    //if (!raw) return false;
    console.log("💒1");
    var ALL = "All Infra 1 tenders";
    var NOT_REL = "Not relevant to additional tenders";
    // אם זה בדיוק אחד משני הערכים המיוחדים
    if (allowAllInfra && raw === ALL)
        return true;
    console.log("💒2");
    if (allowNotRelevant && raw === NOT_REL)
        return true;
    console.log("💒3");
    if (raw === '  ')
        return true;
    console.log("💒4");
    // אחרת: חייב להיות CSV של מכרזים קיימים בלבד
    var parts = raw
        .split(",")
        .map(function (s) { return s.trim(); })
        .filter(Boolean);
    console.log("💒5");
    console.log("💒6");
    // חובה שכל חלק יהיה מכרז קיים
    if (parts.length != 0) {
        for (var _i = 0, parts_1 = parts; _i < parts_1.length; _i++) {
            var p = parts_1[_i];
            console.log("💒7 p ", p);
            if (!validTitles.has(p))
                return false;
            console.log("💒8 p ", p);
        }
    }
    console.log("💒9");
    return true;
}
function sleep(ms) {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            return [2 /*return*/, new Promise(function (res) { return setTimeout(res, ms); })];
        });
    });
}
export function buildPmoUpdatePayloadFromItem(pmoItem) {
    // Fields you should NOT update back to SharePoint
    var BLOCK = new Set([
        "Id",
        "GUID",
        "Created",
        "Modified",
        "Author",
        "Editor",
        "AuthorId",
        "EditorId",
        "Attachments",
        "AttachmentFiles",
        "FileRef",
        "FileLeafRef",
        "odata.type",
        "__metadata",
        "Integration",
    ]);
    var out = {};
    for (var _i = 0, _a = Object.entries(pmoItem !== null && pmoItem !== void 0 ? pmoItem : {}); _i < _a.length; _i++) {
        var _b = _a[_i], k = _b[0], v = _b[1];
        if (BLOCK.has(k))
            continue;
        // reuse your helper: normalizeForSp
        var nv = normalizeForSp(v);
        if (nv !== undefined)
            out[k] = nv;
    }
    console.log("out ", out);
    return out;
}
export function updateAutoCreatedPmoDecisionItem(params) {
    return __awaiter(this, void 0, void 0, function () {
        var sp, integrationId, pmoItem, _a, pmoLinkFieldInternalName, _b, pmoListID, pmoList, linkIdField, maxTries, i, one, two, found, pmoDecisionId, updatePayload;
        return __generator(this, function (_c) {
            switch (_c.label) {
                case 0:
                    sp = params.sp, integrationId = params.integrationId, pmoItem = params.pmoItem, _a = params.pmoLinkFieldInternalName, pmoLinkFieldInternalName = _a === void 0 ? "Integration" : _a, _b = params.pmoListID, pmoListID = _b === void 0 ? "e5e8eaea-16db-49d3-ad7c-62f5a2bdd97a" : _b;
                    console.log("in updateAutoCreatedPmoDecisionItem");
                    pmoList = sp.web.lists.getById(pmoListID);
                    linkIdField = pmoLinkFieldInternalName.endsWith("Id")
                        ? pmoLinkFieldInternalName
                        : "".concat(pmoLinkFieldInternalName, "Id");
                    maxTries = 30;
                    i = 0;
                    _c.label = 1;
                case 1:
                    if (!(i < maxTries)) return [3 /*break*/, 10];
                    console.log("💄");
                    return [4 /*yield*/, pmoList.items
                            .select("*")()];
                case 2:
                    one = _c.sent();
                    console.log(" integrationId ", integrationId, " 1 one ", one);
                    return [4 /*yield*/, pmoList.items
                            .filter("IntegrationId eq ".concat(integrationId))()];
                case 3:
                    two = _c.sent();
                    console.log("2 two ", two);
                    return [4 /*yield*/, pmoList.items
                            .filter("IntegrationId eq ".concat(integrationId))
                            .select("Id")
                            .top(1)()];
                case 4:
                    found = _c.sent();
                    console.log("🤢");
                    if (!(found === null || found === void 0 ? void 0 : found.length)) return [3 /*break*/, 7];
                    console.log("🥳");
                    pmoDecisionId = found[0].Id;
                    updatePayload = buildPmoUpdatePayloadFromItem(pmoItem);
                    console.log("🎆 updatePayload ", updatePayload);
                    return [4 /*yield*/, pmoList.items.getById(pmoListID).update(updatePayload)];
                case 5:
                    _c.sent();
                    console.log("🎆 pmoDecisionId ", pmoDecisionId);
                    return [4 /*yield*/, pmoList.items.getById(pmoDecisionId).update(updatePayload)];
                case 6:
                    _c.sent();
                    console.log("🎆");
                    return [2 /*return*/, { ok: true, pmoDecisionId: pmoDecisionId }];
                case 7: return [4 /*yield*/, sleep(700)];
                case 8:
                    _c.sent();
                    _c.label = 9;
                case 9:
                    i++;
                    return [3 /*break*/, 1];
                case 10: throw new Error("\u05DC\u05D0 \u05E0\u05DE\u05E6\u05D0 \u05E4\u05E8\u05D9\u05D8 \u05D0\u05D5\u05D8\u05D5\u05DE\u05D8\u05D9 \u05D1-".concat(pmoListID, " \u05E2\u05D1\u05D5\u05E8 IntegrationId=").concat(integrationId, ". \u05D1\u05D3\u05E7\u05D9 \u05E9\u05D4\u05D0\u05D5\u05D8\u05D5\u05DE\u05E6\u05D9\u05D4 \u05DE\u05DE\u05DC\u05D0\u05EA \u05D0\u05EA \u05D4\u05E9\u05D3\u05D4 ").concat(linkIdField, "."));
            }
        });
    });
}
function normalizePayloadForSpAdd(payload) {
    var out = {};
    for (var _i = 0, _a = Object.entries(payload); _i < _a.length; _i++) {
        var _b = _a[_i], k = _b[0], v = _b[1];
        // null/undefined נשארים
        if (v == null) {
            out[k] = v;
            continue;
        }
        // אם הגיע בצורה { results: [...] } → להמיר למערך [...]
        if (typeof v === "object" && !Array.isArray(v) && Array.isArray(v.results)) {
            out[k] = v.results;
            continue;
        }
        out[k] = v;
    }
    return out;
}
function normalizeForSp(v) {
    if (v === undefined)
        return undefined; // לא שולחים בכלל
    if (v === null)
        return null; // שולחים null
    if (Array.isArray(v))
        return { results: v }; // MultiChoice/MultiText
    if (typeof v === "object" && (v === null || v === void 0 ? void 0 : v.results))
        return v; // כבר בפורמט results
    return v;
}
export function buildIntegrationPayloadFromPmo(integrationItem, pmoItem, pmoToIntegrationMap, extra) {
    if (extra === void 0) { extra = {}; }
    var payload = {};
    for (var _i = 0, _a = Object.entries(integrationItem !== null && integrationItem !== void 0 ? integrationItem : {}); _i < _a.length; _i++) {
        var _b = _a[_i], k = _b[0], v = _b[1];
        var nv = normalizeForSp(v);
        if (nv !== undefined)
            payload[k] = nv;
    }
    for (var _c = 0, _d = Object.entries(pmoToIntegrationMap); _c < _d.length; _c++) {
        var _e = _d[_c], pmoInternal = _e[0], integrationInternal = _e[1];
        var v = pmoItem === null || pmoItem === void 0 ? void 0 : pmoItem[pmoInternal];
        var nv = normalizeForSp(v);
        if (integrationItem.TenderPhase.indexOf("Phase 1") === -1) { //===
            if (pmoInternal === 'SubCategory') {
                nv = normalizeForSp(integrationItem === null || integrationItem === void 0 ? void 0 : integrationItem[integrationInternal]);
            }
            if (pmoInternal === 'Assignedto') {
                nv = normalizeForSp(integrationItem === null || integrationItem === void 0 ? void 0 : integrationItem[integrationInternal]);
            }
        }
        if (pmoInternal != 'SubCategory') {
            if (nv !== undefined)
                payload[integrationInternal] = nv;
        }
        if (pmoInternal === 'SubCategory') {
            if (nv != undefined && nv != null) {
                payload[integrationInternal] = nv;
            }
            else {
                delete payload[integrationInternal];
            }
        }
    }
    payload['DecisionAppliesToOtherWorksTende0'] = true;
    // תוספות/override (למשל TenderNumber לכל פריט חדש)
    for (var _f = 0, _g = Object.entries(extra); _f < _g.length; _f++) {
        var _h = _g[_f], k = _h[0], v = _h[1];
        var nv = normalizeForSp(v);
        if (nv !== undefined)
            payload[k] = nv;
    }
    return payload;
}
export function splitTenderAndCreateIntegrationItems(params) {
    var _a, _b;
    return __awaiter(this, void 0, void 0, function () {
        var sp, integrationListId, workTendersListId, pmoItem, itegrationItem, tenderSourceInternalName, _c, integrationTenderInternalName, pmoToIntegrationMap, linkFieldInternalName, linkValue, _d, workTenderTitleField, _e, workTenderOlmField, _f, integrationOlmField, _g, decisionAppliesFieldInternalName, _h, 
        // 🆕 PMO clone settings
        pmoDecisionsListId, _j, pmoIntegrationLookupIdField, _k, pmoDecisionAppliesFieldInternalName, _l, pmoSentProtocolFieldInternalName, raw, worktendersList, tenders, tendersByTitle, selectedTitles, phaseToExclude_1, list, pmoList, _i, selectedTitles_1, tenderTitle, wt, olmFromWorkTender, extra, payload, fixed, addRes, createdIntegration, pmoClone, pmoFixed;
        var _m;
        return __generator(this, function (_o) {
            switch (_o.label) {
                case 0:
                    sp = params.sp, integrationListId = params.integrationListId, workTendersListId = params.workTendersListId, pmoItem = params.pmoItem, itegrationItem = params.itegrationItem, tenderSourceInternalName = params.tenderSourceInternalName, _c = params.integrationTenderInternalName, integrationTenderInternalName = _c === void 0 ? "TenderNumber" : _c, pmoToIntegrationMap = params.pmoToIntegrationMap, linkFieldInternalName = params.linkFieldInternalName, linkValue = params.linkValue, _d = params.workTenderTitleField, workTenderTitleField = _d === void 0 ? "Title" : _d, _e = params.workTenderOlmField, workTenderOlmField = _e === void 0 ? "OriginatingLineManager" : _e, _f = params.integrationOlmField, integrationOlmField = _f === void 0 ? "OriginatingLineManager" : _f, _g = params.decisionAppliesFieldInternalName, decisionAppliesFieldInternalName = _g === void 0 ? "DecisionappliestootherWorksTende" : _g, _h = params.pmoDecisionsListId, pmoDecisionsListId = _h === void 0 ? 'e5e8eaea-16db-49d3-ad7c-62f5a2bdd97a' : _h, _j = params.pmoIntegrationLookupIdField, pmoIntegrationLookupIdField = _j === void 0 ? "IntegrationId" : _j, _k = params.pmoDecisionAppliesFieldInternalName, pmoDecisionAppliesFieldInternalName = _k === void 0 ? "DecisionAppliesToOtherWorksTende" : _k, _l = params.pmoSentProtocolFieldInternalName, pmoSentProtocolFieldInternalName = _l === void 0 ? "sentProtocol" : _l;
                    console.log("itegrationItem ", itegrationItem);
                    console.log("pmoItem ", pmoItem);
                    console.log("🪁 pmoItem.sentProtocol ", pmoItem.sentProtocol);
                    if (itegrationItem.coppiedFrom != null) {
                        return [2 /*return*/];
                    }
                    raw = String((_a = pmoItem === null || pmoItem === void 0 ? void 0 : pmoItem[tenderSourceInternalName]) !== null && _a !== void 0 ? _a : "").trim();
                    console.log("🩰 raw ", raw);
                    if (!raw)
                        return [2 /*return*/];
                    console.log("🩰 1");
                    if (raw === "Not relevant to additional tenders") {
                        console.log("raw is not relevant ", raw);
                        return [2 /*return*/];
                    }
                    console.log("🩰 2");
                    worktendersList = sp.web.lists.getById(workTendersListId);
                    console.log("🩰 3");
                    return [4 /*yield*/, worktendersList.items
                            .select(workTenderTitleField, workTenderOlmField)()];
                case 1:
                    tenders = _o.sent();
                    console.log("🩰 4");
                    console.log("tenders ", tenders);
                    tendersByTitle = new Map(tenders.map(function (t) { var _a; return [String((_a = t === null || t === void 0 ? void 0 : t[workTenderTitleField]) !== null && _a !== void 0 ? _a : "").trim(), t]; }));
                    console.log("🩰 5");
                    selectedTitles = [];
                    console.log("🩰 6");
                    if (raw === "All Infra 1 tenders") {
                        selectedTitles = tenders
                            .map(function (t) { var _a; return String((_a = t === null || t === void 0 ? void 0 : t[workTenderTitleField]) !== null && _a !== void 0 ? _a : "").trim(); })
                            .filter(Boolean);
                        console.log("🩰 7");
                        phaseToExclude_1 = String((_b = itegrationItem === null || itegrationItem === void 0 ? void 0 : itegrationItem.TenderNumber) !== null && _b !== void 0 ? _b : "").trim();
                        console.log("phaseToExclude ", phaseToExclude_1);
                        console.log("🩰 8");
                        selectedTitles = selectedTitles.filter(function (title) { return title !== phaseToExclude_1; });
                        console.log("🩰 9");
                        console.log("raw === 'All Infra 1 tenders' -> selectedTitles ", selectedTitles);
                    }
                    else {
                        console.log("🩰 10");
                        selectedTitles = raw
                            .split(",")
                            .map(function (s) { return s.trim(); })
                            .filter(Boolean);
                        console.log("🩰 11");
                        console.log("raw != 'All Infra 1 tenders' -> selectedTitles ", selectedTitles);
                    }
                    if (selectedTitles.length === 0)
                        return [2 /*return*/];
                    console.log("🩰 12");
                    list = sp.web.lists.getById(integrationListId);
                    console.log("🩰 13");
                    pmoList = pmoDecisionsListId ? sp.web.lists.getById(pmoDecisionsListId) : null;
                    console.log("🩰 14");
                    _i = 0, selectedTitles_1 = selectedTitles;
                    _o.label = 2;
                case 2:
                    if (!(_i < selectedTitles_1.length)) return [3 /*break*/, 6];
                    tenderTitle = selectedTitles_1[_i];
                    console.log("💒tenderTitle- ", tenderTitle, "  itegrationItem.TenderNumber -", itegrationItem.TenderNumber);
                    if (tenderTitle === itegrationItem.TenderPhase)
                        return [3 /*break*/, 5];
                    if (tenderTitle.trim() === itegrationItem.TenderNumber.trim())
                        return [3 /*break*/, 5];
                    wt = tendersByTitle.get(String(tenderTitle).trim());
                    olmFromWorkTender = wt ? wt === null || wt === void 0 ? void 0 : wt[workTenderOlmField] : undefined;
                    extra = (_m = {},
                        _m[integrationTenderInternalName] = tenderTitle,
                        _m);
                    // ✅ זה השדה שאת רוצה שיהיה "Not relevant..." (נשמר כ-MultiChoice)
                    extra[decisionAppliesFieldInternalName] = {
                        results: ["Not relevant to additional tenders"],
                    };
                    extra["coppiedFrom"] = String(itegrationItem.NTA_x2019_s_x0020_reference);
                    extra["LM_x2019_sreference"] = String(itegrationItem.NTA_x2019_s_x0020_reference);
                    // ✅ אם מצאנו OLM ב-WorkTenders – נשפוך אותו ל-Integration
                    if (itegrationItem.Category === null || itegrationItem.Category === undefined) {
                        extra["Category"] = [];
                    }
                    if (olmFromWorkTender !== undefined && olmFromWorkTender !== null) {
                        extra[integrationOlmField] = olmFromWorkTender;
                    }
                    // lookup/reference אם צריך
                    if (linkFieldInternalName && linkValue != null) {
                        extra[linkFieldInternalName] = linkValue;
                    }
                    payload = buildIntegrationPayloadFromPmo(itegrationItem, pmoItem, pmoToIntegrationMap, extra);
                    console.log("🏔️ integration addPayload", payload);
                    fixed = normalizePayloadForSpAdd(payload);
                    console.log("✅ fixed payload", fixed);
                    return [4 /*yield*/, list.items.add(fixed)];
                case 3:
                    addRes = _o.sent();
                    console.log("addRes =", addRes);
                    console.log("keys(addRes) =", addRes ? Object.keys(addRes) : null);
                    createdIntegration = addRes;
                    console.log("✅ createdIntegration:", createdIntegration);
                    if (!(pmoList && (createdIntegration === null || createdIntegration === void 0 ? void 0 : createdIntegration.Id))) return [3 /*break*/, 5];
                    pmoClone = __assign({}, (pmoItem || {}));
                    // ניקוי שדות מערכת/זהויות כדי שלא יפילו add
                    delete pmoClone.Id;
                    delete pmoClone.ID;
                    delete pmoClone.odata;
                    delete pmoClone["odata.type"];
                    delete pmoClone["odata.id"];
                    delete pmoClone["odata.etag"];
                    delete pmoClone["odata.editLink"];
                    delete pmoClone.AuthorId;
                    delete pmoClone.EditorId;
                    delete pmoClone.Created;
                    delete pmoClone.Modified;
                    delete pmoClone["odata.metadata"];
                    if (itegrationItem.SubCategory === null || itegrationItem.SubCategory === undefined) {
                        delete pmoClone["SubCategory"];
                    }
                    if (itegrationItem.Category === null || itegrationItem.Category === undefined) {
                        delete pmoClone["Category"];
                    }
                    // ✅ חריגים שביקשת:
                    // Lookup ל־Integration החדש (בד"כ זה IntegrationId)
                    pmoClone[pmoIntegrationLookupIdField] = createdIntegration.Id;
                    // MultiChoice / choice-like לפי מה שביקשת
                    pmoClone[pmoDecisionAppliesFieldInternalName] = {
                        results: ["Not relevant to additional tenders"],
                    };
                    // sentProtocol = false
                    pmoClone[pmoSentProtocolFieldInternalName] = pmoItem.sentProtocol; //false;
                    pmoFixed = normalizePayloadForSpAdd(pmoClone);
                    console.log("🧾 PMO clone payload:", pmoFixed);
                    return [4 /*yield*/, pmoList.items.add(pmoFixed)];
                case 4:
                    _o.sent();
                    console.log("✅ PMO clone item created for integration Id:", createdIntegration.Id);
                    _o.label = 5;
                case 5:
                    _i++;
                    return [3 /*break*/, 2];
                case 6: return [2 /*return*/];
            }
        });
    });
}
function splitCommaList(raw) {
    return raw
        .split(",")
        .map(function (s) { return s.trim(); })
        .filter(Boolean);
}
function isEmptyValue(v) {
    if (v === null || v === undefined)
        return true;
    if (typeof v === "string")
        return v.trim() === "";
    if (Array.isArray(v))
        return v.length === 0;
    return false;
}
function isSystemFieldLocal(internal) {
    var s = internal.toLowerCase();
    return (s === "id" ||
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
        s.startsWith("odata__colortag"));
}
function pickValue(primary, secondary, internal) {
    var a = primary === null || primary === void 0 ? void 0 : primary[internal];
    if (!isEmptyValue(a))
        return a;
    var b = secondary === null || secondary === void 0 ? void 0 : secondary[internal];
    if (!isEmptyValue(b))
        return b;
    return undefined;
}
/**
 * הופך ערכים לפורמט ש-SharePoint REST מצפה לו במקרים נפוצים.
 * (הכי חשוב: MultiChoice => { results: [...] })
 */
function coerceValueForSp(typeAsString, val) {
    var t = (typeAsString || "").toLowerCase();
    // תאריך
    if (t.includes("date") || t.includes("datetime")) {
        if (val instanceof Date)
            return val.toISOString();
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
/////////////////////////////////////////////////////////////////////////////////
function buildMergedClonePayload(primary, secondary, fieldInfoMap, overrides) {
    var payload = {};
    for (var _i = 0, _a = Object.keys(fieldInfoMap); _i < _a.length; _i++) {
        var internal = _a[_i];
        var info = fieldInfoMap[internal];
        if (!info)
            continue;
        if (info.Hidden || info.ReadOnlyField || info.Sealed)
            continue;
        if (isSystemFieldLocal(internal))
            continue;
        // לוקחים ערך מהראשון, ואם ריק אז מהשני
        var v = pickValue(primary, secondary, internal);
        if (v === undefined)
            continue;
        payload[internal] = coerceValueForSp(info.TypeAsString || "", v);
    }
    // overrides (למשל TenderNumber = X)
    for (var _b = 0, _c = Object.keys(overrides); _b < _c.length; _b++) {
        var k = _c[_b];
        payload[k] = overrides[k];
    }
    return payload;
}
export function splitTenderAndCreateItemsFromTwoSources(params) {
    var _a, _b;
    return __awaiter(this, void 0, void 0, function () {
        var sp, listId, primaryItem, secondaryItem, fieldInfoMap, _c, tenderFieldInternalName, raw, parts, list, _i, _d, x, addPayload;
        var _e;
        return __generator(this, function (_f) {
            switch (_f.label) {
                case 0:
                    sp = params.sp, listId = params.listId, primaryItem = params.primaryItem, secondaryItem = params.secondaryItem, fieldInfoMap = params.fieldInfoMap, _c = params.tenderFieldInternalName, tenderFieldInternalName = _c === void 0 ? "DecisionAppliesToOtherWorksTende" : _c;
                    console.log("🏔️🏔️🏔️ primaryItem ", primaryItem);
                    console.log("🏔️🏔️🏔️ secondaryItem ", secondaryItem);
                    raw = String((_b = (_a = primaryItem === null || primaryItem === void 0 ? void 0 : primaryItem[tenderFieldInternalName]) !== null && _a !== void 0 ? _a : secondaryItem === null || secondaryItem === void 0 ? void 0 : secondaryItem[tenderFieldInternalName]) !== null && _b !== void 0 ? _b : "").trim();
                    console.log("🏔️🏔️🏔️raw ", raw);
                    parts = splitCommaList(raw);
                    if (parts.length <= 1)
                        return [2 /*return*/];
                    console.log("🏔️🏔️🏔️parts ", parts);
                    list = sp.web.lists.getById(listId);
                    _i = 0, _d = parts.slice(0);
                    _f.label = 1;
                case 1:
                    if (!(_i < _d.length)) return [3 /*break*/, 4];
                    x = _d[_i];
                    addPayload = buildMergedClonePayload(primaryItem, secondaryItem, fieldInfoMap, (_e = {}, _e['TenderNumber'] = x, _e));
                    console.log("🏔️🏔️🏔️addPayload ", addPayload);
                    return [4 /*yield*/, list.items.add(addPayload)];
                case 2:
                    _f.sent();
                    _f.label = 3;
                case 3:
                    _i++;
                    return [3 /*break*/, 1];
                case 4: return [2 /*return*/];
            }
        });
    });
}
// מחזיר: ["M3"] או ["M2","M3"] או ["ALL"] או שילובים
export function myOLM(sp, generalRoleDefinitionListId, email) {
    return __awaiter(this, void 0, void 0, function () {
        var userEmail, ROLE_COLUMNS, ALL_COLUMNS, list, rows, roles, emailInTextField, _loop_1, _i, rows_1, row, result;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    console.log("🎍 in my OLM");
                    userEmail = String(email || "").trim().toLowerCase();
                    if (!userEmail)
                        return [2 /*return*/, []];
                    console.log("🎍 hase email ");
                    ROLE_COLUMNS = {
                        M1: ["M1TenderTeam", "LawyerM1TenderTeam"],
                        M2: ["M2TenderTeam", "LawyerM2TenderTeam"],
                        M3: ["M3TenderTeam", "LawyerM3TenderTeam"],
                    };
                    ALL_COLUMNS = ["IntegrationTeamLawyer", "IntegrationTeam", "FinancialAdvisorIntegrationTeam"];
                    console.log("🎍 1 ");
                    //const select = ["Id", ...peopleCols].join(",");
                    console.log("🎍 2 ");
                    list = sp.web.lists.getById(generalRoleDefinitionListId);
                    console.log("🎍 3 ");
                    return [4 /*yield*/, list.items
                            .select()()];
                case 1:
                    rows = _a.sent();
                    console.log("🎍 4 , rows ", rows);
                    if (!(rows === null || rows === void 0 ? void 0 : rows.length))
                        return [2 /*return*/, []];
                    console.log("🎍 5 ");
                    roles = new Set();
                    console.log("🎍 6 ");
                    emailInTextField = function (fieldVal) {
                        if (!fieldVal)
                            return false;
                        return String(fieldVal).trim().toLowerCase() === userEmail;
                    };
                    console.log("🎍 7 ");
                    _loop_1 = function (row) {
                        // ALL
                        for (var _b = 0, ALL_COLUMNS_1 = ALL_COLUMNS; _b < ALL_COLUMNS_1.length; _b++) {
                            var col = ALL_COLUMNS_1[_b];
                            if (emailInTextField(row[col])) {
                                roles.add("ALL");
                            }
                        }
                        // M1 / M2 / M3
                        Object.keys(ROLE_COLUMNS).forEach(function (role) {
                            for (var _i = 0, _a = ROLE_COLUMNS[role]; _i < _a.length; _i++) {
                                var col = _a[_i];
                                if (emailInTextField(row[col])) {
                                    console.log("emailInTextField(row[col]  ", row[col]);
                                    roles.add(role);
                                }
                            }
                        });
                    };
                    for (_i = 0, rows_1 = rows; _i < rows_1.length; _i++) {
                        row = rows_1[_i];
                        _loop_1(row);
                    }
                    console.log("🎍 8 ");
                    result = Array.from(roles);
                    console.log("🎍 roles:", result);
                    return [2 /*return*/, result];
            }
        });
    });
}
function ensureReloadGuardList(sp) {
    return __awaiter(this, void 0, void 0, function () {
        var list, err_1, res;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    console.log("🕐 Starting ensureReloadGuardList");
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, , 5]);
                    // ⚠️ CRITICAL: Ensure sp.web is actually initialized
                    if (!(sp === null || sp === void 0 ? void 0 : sp.web)) {
                        throw new Error("SP context not initialized");
                    }
                    // Test if list exists
                    list = sp.web.lists.getByTitle(RELOAD_GUARD_LIST_TITLE);
                    return [4 /*yield*/, list.select("Id")()];
                case 2:
                    _a.sent(); // Lightweight check
                    console.log("✅ List exists");
                    return [3 /*break*/, 5];
                case 3:
                    err_1 = _a.sent();
                    console.log("📝 List doesn't exist, creating...", err_1.message);
                    return [4 /*yield*/, sp.web.lists.add(RELOAD_GUARD_LIST_TITLE, "", 100, false)];
                case 4:
                    res = _a.sent();
                    list = res.list;
                    return [3 /*break*/, 5];
                case 5: 
                // Ensure fields exist...
                // (rest of your field creation code)
                return [2 /*return*/, list];
            }
        });
    });
}
function getOrCreateReloadGuardItem(sp, list) {
    return __awaiter(this, void 0, void 0, function () {
        var me, items, addRes;
        var _a, _b;
        return __generator(this, function (_c) {
            switch (_c.label) {
                case 0: return [4 /*yield*/, sp.web.currentUser()];
                case 1:
                    me = _c.sent();
                    return [4 /*yield*/, list.items
                            .filter("".concat(RELOAD_GUARD_USER_FIELD, "/Id eq ").concat(me.Id))
                            .select("Id", RELOAD_GUARD_FLAG_FIELD, "".concat(RELOAD_GUARD_USER_FIELD, "/Id"))
                            .expand(RELOAD_GUARD_USER_FIELD)()];
                case 2:
                    items = _c.sent();
                    if (items.length > 0)
                        return [2 /*return*/, items[0]];
                    return [4 /*yield*/, list.items.add((_a = {},
                            _a["".concat(RELOAD_GUARD_USER_FIELD, "Id")] = me.Id,
                            _a[RELOAD_GUARD_FLAG_FIELD] = false,
                            _a))];
                case 3:
                    addRes = _c.sent();
                    return [2 /*return*/, (_b = { Id: addRes.data.Id }, _b[RELOAD_GUARD_FLAG_FIELD] = false, _b)];
            }
        });
    });
}
function markReloadedAndReload(sp) {
    return __awaiter(this, void 0, void 0, function () {
        var list, item, alreadyReloaded;
        var _a;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0: return [4 /*yield*/, ensureReloadGuardList(sp)];
                case 1:
                    list = _b.sent();
                    return [4 /*yield*/, getOrCreateReloadGuardItem(sp, list)];
                case 2:
                    item = _b.sent();
                    alreadyReloaded = !!(item === null || item === void 0 ? void 0 : item[RELOAD_GUARD_FLAG_FIELD]);
                    if (!!alreadyReloaded) return [3 /*break*/, 4];
                    return [4 /*yield*/, list.items.getById(item.Id).update((_a = {},
                            _a[RELOAD_GUARD_FLAG_FIELD] = true,
                            _a))];
                case 3:
                    _b.sent();
                    window.location.reload();
                    return [2 /*return*/, true]; // did reload
                case 4: return [2 /*return*/, false]; // no reload
            }
        });
    });
}
function resetReloadGuard(sp) {
    return __awaiter(this, void 0, void 0, function () {
        var list, me, items, _a;
        var _b;
        return __generator(this, function (_c) {
            switch (_c.label) {
                case 0:
                    _c.trys.push([0, 6, , 7]);
                    return [4 /*yield*/, ensureReloadGuardList(sp)];
                case 1:
                    list = _c.sent();
                    return [4 /*yield*/, sp.web.currentUser()];
                case 2:
                    me = _c.sent();
                    return [4 /*yield*/, list.items
                            .filter("".concat(RELOAD_GUARD_USER_FIELD, "/Id eq ").concat(me.Id))
                            .select("Id", RELOAD_GUARD_FLAG_FIELD, "".concat(RELOAD_GUARD_USER_FIELD, "/Id"))
                            .expand(RELOAD_GUARD_USER_FIELD)()];
                case 3:
                    items = _c.sent();
                    if (!(items.length > 0)) return [3 /*break*/, 5];
                    return [4 /*yield*/, list.items.getById(items[0].Id).update((_b = {},
                            _b[RELOAD_GUARD_FLAG_FIELD] = false,
                            _b))];
                case 4:
                    _c.sent();
                    _c.label = 5;
                case 5: return [3 /*break*/, 7];
                case 6:
                    _a = _c.sent();
                    return [3 /*break*/, 7];
                case 7: return [2 /*return*/];
            }
        });
    });
}
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
function normalizeStepsConfigToInternal(steps, titleToInternal) {
    var out = {};
    var stepNames = Object.keys(steps || {});
    for (var s = 0; s < stepNames.length; s++) {
        var step = stepNames[s];
        var arr = steps[step] || [];
        var internals = [];
        var seen = {};
        for (var i = 0; i < arr.length; i++) {
            var name_1 = arr[i];
            if (!name_1)
                continue;
            var internal = titleToInternal[name_1] || name_1;
            if (internal && !seen[internal]) {
                internals.push(internal);
                seen[internal] = true;
            }
        }
        out[step] = internals;
    }
    return out;
}
function getPhaseViewKeyFromTenderPhase(tenderPhaseRaw) {
    var s = String(tenderPhaseRaw || '').trim().toLowerCase();
    if (!s)
        return '';
    // אם מתחיל ב־Phase 1 / Phase 2 / Phase 3 – ניקח את ה־prefix
    if (s.indexOf('phase 1') != -1)
        return 'Phase 1';
    if (s.indexOf('phase 2') != -1)
        return 'Phase 2';
    if (s.indexOf('phase 3') != -1)
        return 'Phase 3';
    // ברירת מחדל – אין סינון לפי תצוגה
    return '';
}
var FormApp = function (_a) {
    var _b;
    var context = _a.context, pmoListTitle = _a.pmoListTitle, _c = _a.pmoIntegrationLookupName, pmoIntegrationLookupName = _c === void 0 ? 'Integration' : _c, _d = _a.stepsConfig, stepsConfig = _d === void 0 ? stepsConfigJson : _d, _e = _a.isEditMode, isEditMode = _e === void 0 ? false : _e;
    var PMO_LIST_ID = "e5e8eaea-16db-49d3-ad7c-62f5a2bdd97a";
    var STATUS_DECISION_REACHED = "Decision reached - To be included in Protocol";
    // אם זה שם פנימי שונה אצלכם - תעדכני פה:
    var INTEGRATION_TEAM_STATUS_FIELD = "INTEGRATIONTEAMSTATUS";
    var PMO_SENT_PROTOCOL_FIELD = "sentProtocol";
    var sp = useMemo(function () { return getSP(context); }, [context]);
    var _f = useState(null), me = _f[0], setMe = _f[1];
    var _g = useState([]), integrationChoices = _g[0], setIntegrationChoices = _g[1];
    var _h = useState(null), integrationId = _h[0], setIntegrationId = _h[1];
    // const [comboOpen, setComboOpen] = useState<boolean>(false);
    var _j = useState('ALL'), olmFilter = _j[0], setOlmFilter = _j[1];
    // Integration (תצוגה בלבד) + תוויות + מטא
    var _k = useState(null), integrationItem = _k[0], setIntegrationItem = _k[1];
    var _l = useState({}), integrationFieldInfoMap = _l[0], setIntegrationFieldInfoMap = _l[1];
    var _m = useState({}), integrationLabels = _m[0], setIntegrationLabels = _m[1];
    // סדר שדות לתצוגה לפי View של רשימת Integration
    var _o = useState([]), integrationViewFieldOrder = _o[0], setIntegrationViewFieldOrder = _o[1];
    //const [, setIntegrationViewName] = useState<string>('');
    var _p = useState(''), integrationSearch = _p[0], setIntegrationSearch = _p[1];
    // PMO (עריכה) + תוויות + מטא
    var _q = useState(null), pmoItem = _q[0], setPmoItem = _q[1];
    var _r = useState({}), pmoDraft = _r[0], setPmoDraft = _r[1];
    // placeholder לשדות ספציפיים כשיש Accept
    var placeholders = React.useMemo(function () {
        var decision = String((pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.DecisionRegardingProposedChange) || '').trim();
        if (decision === 'Accept' || decision === 'Partially accepted') {
            var msg_1 = 'Enter final wording for publication here';
            return {
                RevisedWordingFinalForPublicatio: msg_1,
                dog: msg_1
            };
        }
        return {};
    }, [pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.DecisionRegardingProposedChange]);
    var _s = React.useState(false), isSplitting = _s[0], setIsSplitting = _s[1];
    var _t = useState({}), pmoLabels = _t[0], setPmoLabels = _t[1];
    var _u = useState({}), pmoFieldInfoMap = _u[0], setPmoFieldInfoMap = _u[1];
    // שלבים
    var _v = useState({}), stepsInternal = _v[0], setStepsInternal = _v[1];
    var _w = useState(''), activeStep = _w[0], setActiveStep = _w[1];
    // שדות לפי תצוגת SharePoint (View) לפי TenderPhase
    var _x = useState([]), viewFieldOrder = _x[0], setViewFieldOrder = _x[1];
    var _y = useState(''), currentViewName = _y[0], setCurrentViewName = _y[1];
    // הרשאות שדה
    var _z = useState({}), fieldPermMap = _z[0], setFieldPermMap = _z[1];
    var _0 = React.useState([]), myRoles = _0[0], setMyRoles = _0[1];
    var _1 = useState({
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
    }), roleUsers = _1[0], setRoleUsers = _1[1];
    // UI
    var _2 = useState(false), busy = _2[0], setBusy = _2[1];
    var _3 = useState(null), msg = _3[0], setMsg = _3[1];
    var _4 = React.useState(new Set()), allowedTenderTitles = _4[0], setAllowedTenderTitles = _4[1];
    var _5 = React.useState(false), allowedTenderTitlesReady = _5[0], setAllowedTenderTitlesReady = _5[1];
    // ----- שדות חובה בטופס PMO decisions -----
    var REQUIRED_FIELDS = [
        'DecisionRegardingProposedChange',
        'DecisionAppliesToOtherWorksTende',
        'RevisionIncludesChangeInTenderDo',
        'DecisionDate'
    ];
    var TENDER_TEAM_FIELDS = [
        "StatusOfRFCresponseOrTcRFC", //
        "RFCorTcRFCasPublishedByNTaToBeFi", //
        "RFCresponseAsPublishedToBeFilled", //
        "Addendum", //
        "addendumDate",
        "RFCResponseLetterNo", //
        "TenderCommitteeApprovalDate", //
        "RevisedWordingFinalForPublicatio", //
        "IntegrationTeamDecisionImplement", //
        'DueDateCalculated',
        'ActualDate',
        'SubCategory',
        'Assignedto'
    ];
    // 🆕 מפת תוויות קשיחה – fallback במקרה שאין internalToTitle מה-SharePoint
    var PMO_LABEL_OVERRIDES = {
        Integration: 'NTA reference#',
        DecisionRegardingProposedChange: 'Decision Regarding Proposed Change',
        DecisionRegardingProposedChangeC: 'Decision Regarding Proposed Change - comments',
        DecisionAppliesToOtherWorksTende: 'Decision applies to other Works Tenders?',
        IntegrationTeamDecisionImplement: 'Decision Implementation Status',
        RevisionIncludesChangeInTenderDo: 'Revision includes change in Tender Documents? (Y/N) If Y, Addendum required',
        RFCResponseLetterNo: 'RFC response Letter no.',
        RFCresponseAsPublishedToBeFilled: 'RFC response as published  (To be filled in after publication)',
        Addendum: 'Addendum #',
        addendumDate: 'Addendum Date',
        TenderCommitteeApprovalDate: 'Tender Committee Approval Date',
        StatusOfRFCresponseOrTcRFC: 'Status Of RFC response Or Tc RFC',
        DecisionDate: 'Decision Date',
        RevisedWordingFinalForPublicatio: 'Revised Wording - final for publication',
        RFCorTcRFCasPublishedByNTaToBeFi: 'RFC / TC RFC as published by NTA (To be filled in after publication)',
        DueDateCalculated: 'Due date (calculated)',
        ActualDate: 'Actual date',
        DateForIntegrationTeamDecisionIm: 'Date For Integration Team Decision Implementation Status',
        INTEGRATIONTEAMSTATUS: 'INTEGRATION TEAM STATUS',
        RFCorTcRFCasPublishedByNTaNew: 'RFCorTcRFCasPublishedByNTaNew',
        sentProtocol: 'sentProtocol',
        Assignedto: 'Assigned to',
        SubCategory: 'Sub - Category'
    };
    // מפה מהיר מ־internalName -> חובה?
    var requiredMap = useMemo(function () {
        var m = {};
        for (var i = 0; i < REQUIRED_FIELDS.length; i++) {
            m[REQUIRED_FIELDS[i]] = true;
        }
        return m;
    }, []);
    // שגיאות ולידציה לשדות (שדה -> טקסט שגיאה)
    var _6 = useState({}), validationErrors = _6[0], setValidationErrors = _6[1];
    // ==== Sync PMO -> INTEGRATION ====
    // מפה בין שמות השדות ב־PMO לבין השדות המקבילים ברשימת INTEGRATION
    // 🟢 אם השמות זהים בשתי הרשימות – פשוט תשאירי את אותו שם גם מימין וגם משמאל.
    // 🟢 תעדכני כאן את כל השדות שאת רוצה לסנכרן.
    var PMO_TO_INTEGRATION_FIELD_MAP = {
        // PMO internal name      : INTEGRATION internal name
        INTEGRATIONTEAMSTATUS: 'INTEGRATIONTEAMSTATUS',
        DecisionRegardingProposedChange: 'DecisionRegardingProposedChange',
        DecisionRegardingProposedChangeC: 'DecisionRegardingProposedChange_',
        DecisionAppliesToOtherWorksTende: 'DecisionappliestootherWorksTende',
        IntegrationTeamDecisionImplement: 'IntegrationTeamDecision',
        RevisionIncludesChangeInTenderDo: 'RevisionincludeschangeinTenderDo',
        RFCResponseLetterNo: 'RFCresponseLetterno',
        RFCresponseAsPublishedToBeFilled: 'RFCresponseaspublished',
        Addendum: 'Addendum',
        addendumDate: 'addendumDate',
        TenderCommitteeApprovalDate: 'TenderCommitteeapprovaldate',
        StatusOfRFCresponseOrTcRFC: 'StatusofRFCresponse_x002f_TCRFC',
        DecisionDate: 'Decisiondate',
        RevisedWordingFinalForPublicatio: 'RevisedWording_x002d_finalforpub',
        dog: 'RFCorTcRFCasPublishedByNTaNew', //'RFC_x002f_TCRFCaspublishedbyNTA_',
        DueDateCalculated: 'Duedate',
        ActualDate: 'Actualdate',
        RFCorTcRFCasPublishedByNTaToBeFi: 'RFC_x002f_TCRFCaspublishedbyNTA_',
        SubCategory: 'Sub_x002d_Category',
        Assignedto: 'Assignedto'
    };
    function syncPmoToIntegration(sp, integrationId, pmoDraft) {
        return __awaiter(this, void 0, void 0, function () {
            var updatePayload, pmoField, integrationField, val, outVal;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        console.log("🫧 saving to integration");
                        if (!integrationId)
                            return [2 /*return*/];
                        updatePayload = {};
                        for (pmoField in PMO_TO_INTEGRATION_FIELD_MAP) {
                            console.log("pmoField ", pmoField);
                            if (!Object.prototype.hasOwnProperty.call(PMO_TO_INTEGRATION_FIELD_MAP, pmoField))
                                continue;
                            console.log("🧸 Object.prototype.hasOwnProperty.call");
                            integrationField = PMO_TO_INTEGRATION_FIELD_MAP[pmoField];
                            val = pmoDraft ? pmoDraft[pmoField] : undefined;
                            if (typeof val === 'undefined')
                                continue;
                            if (val === null)
                                continue;
                            console.log("🐻 not null or undefined ", val);
                            outVal = val;
                            updatePayload[integrationField] = outVal;
                        }
                        if (Object.keys(updatePayload).length === 0) {
                            console.log('syncPmoToIntegration: nothing to update');
                            return [2 /*return*/];
                        }
                        console.log('syncPmoToIntegration →', updatePayload);
                        return [4 /*yield*/, sp.web.lists
                                .getById(INTEGRATION_LIST_ID)
                                .items.getById(integrationId)
                                .update(updatePayload)];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    }
    ////////////////////////////////////////////////////////////////////////////////////
    var filteredIntegrationChoices = useMemo(function () {
        console.log("🦝🦝🦝🦝🦝🦝🦝🦝 filteredIntegrationChoices");
        //const q = integrationSearch.trim().toLowerCase();
        var q = (integrationSearch || '').trim().toLowerCase();
        var arr = integrationChoices;
        return arr.filter(function (opt) {
            var _a;
            var textMatch = !q || String(opt.text || '').toLowerCase().indexOf(q) !== -1;
            var olm = String(((_a = opt.data) === null || _a === void 0 ? void 0 : _a.olm) || '').toUpperCase();
            var olmMatch = olmFilter === 'ALL' || !olmFilter
                ? true
                : olm === olmFilter;
            return textMatch && olmMatch;
        });
    }, [integrationSearch, integrationChoices, olmFilter]);
    var loadIntegrationChoices = function () { return __awaiter(void 0, void 0, void 0, function () {
        var items, myRolesUpper, showAll, opts, i, it, id, raw, plain, oneLine, preview, olm, text;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, sp.web.lists
                        .getById(INTEGRATION_LIST_ID)
                        .items.select('Id', INTEGRATION_PREVIEW_FIELD_INTERNAL, 'NTA_x2019_s_x0020_reference', 'OriginatingLineManager' // 👈 מוסיפים את השדה לפעולת ה־select
                    )
                        .orderBy('Id', false)
                        .top(200)()];
                case 1:
                    items = _a.sent();
                    console.log("loadIntegrationChoices items ", items);
                    myRolesUpper = myRoles.map(function (r) { return r.toUpperCase(); });
                    showAll = myRolesUpper.includes('ALL');
                    opts = [];
                    for (i = 0; i < items.length; i++) {
                        it = items[i];
                        id = it.Id;
                        raw = it[INTEGRATION_PREVIEW_FIELD_INTERNAL];
                        plain = htmlToPlainText(raw);
                        oneLine = plain.replace(/\r?\n/g, ' ').trim();
                        preview = oneLine.length > 30 ? oneLine.substring(0, 30) + '…' : oneLine;
                        olm = String(it.OriginatingLineManager || '').trim().toUpperCase();
                        //const text = preview ? `${id} — ${preview}` : String(id);
                        console.log(preview ? "".concat(id, " \u2014 ").concat(preview) : String(id));
                        text = String(id);
                        if (!showAll && (!olm || !myRolesUpper.includes(olm))) {
                            continue;
                        }
                        opts.push({
                            key: id,
                            text: text,
                            // נשמור את ה-OLM בתוך data כדי שנוכל לסנן אחר כך
                            data: { olm: olm }
                        });
                    }
                    console.log("🥗 opts ", opts);
                    setIntegrationChoices(opts);
                    return [2 /*return*/];
            }
        });
    }); };
    /////////////////
    useEffect(function () {
        if (!sp)
            return;
        if (!myRoles.length)
            return; // ⬅️ זה החלק הקריטי
        loadIntegrationChoices();
    }, [sp, myRoles]);
    React.useEffect(function () {
        var disposed = false;
        (function () { return __awaiter(void 0, void 0, void 0, function () {
            var titles, e_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, fetchAllTenderTitles({
                                sp: sp,
                                workTendersListId: "a33ec5e6-86c0-439a-9eff-f5807b7764d9",
                                titleFieldInternalName: "LinkTitle",
                            })];
                    case 1:
                        titles = _a.sent();
                        console.log("after fetchAllTenderTitles");
                        if (!disposed) {
                            setAllowedTenderTitles(titles);
                            setAllowedTenderTitlesReady(true);
                        }
                        return [3 /*break*/, 3];
                    case 2:
                        e_1 = _a.sent();
                        console.error("Failed to load tenders list titles", e_1);
                        if (!disposed)
                            setAllowedTenderTitlesReady(false);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        }); })();
        return function () { disposed = true; };
    }, [sp]);
    var loadIntegrationMeta = function () { return __awaiter(void 0, void 0, void 0, function () {
        var fields, labels, meta, i, f, internal;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, sp.web.lists
                        .getById(INTEGRATION_LIST_ID)
                        .fields
                        .select('Title', 'InternalName', 'Hidden', 'ReadOnlyField', 'TypeAsString')()];
                case 1:
                    fields = _a.sent();
                    labels = {};
                    meta = {};
                    for (i = 0; i < fields.length; i++) {
                        f = fields[i];
                        internal = f.InternalName;
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
                    return [2 /*return*/];
            }
        });
    }); };
    var loadIntegrationLabels = function () { return __awaiter(void 0, void 0, void 0, function () {
        var fields, map, i, f;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, sp.web.lists.getById(INTEGRATION_LIST_ID).fields.select('Title', 'InternalName')()];
                case 1:
                    fields = _a.sent();
                    map = {};
                    for (i = 0; i < fields.length; i++) {
                        f = fields[i];
                        map[f.InternalName] = f.Title || f.InternalName;
                    }
                    setIntegrationLabels(map);
                    return [2 /*return*/];
            }
        });
    }); };
    var getOriginatingLineManager = function () {
        var raw = String((integrationItem === null || integrationItem === void 0 ? void 0 : integrationItem.OriginatingLineManager) || '').trim().toUpperCase();
        if (raw === 'M1' || raw === 'M2' || raw === 'M3') {
            return raw;
        }
        return '';
    };
    var getIntegrationViewNameForTenderPhase = function (rawPhase) {
        var phase = (rawPhase || '').toLowerCase().trim();
        // כאן תתאימי לשמות התצוגות אצלך ברשימת Integration
        if (phase.indexOf('phase 1') === 0) {
            return 'Phase 1'; // ← שם ה-View ברשימת INTEGRATION
        }
        if (phase.indexOf('phase 2') === 0) {
            return 'Phase 2';
        }
        if (phase.indexOf('phase 3') === 0) {
            return 'Phase 3';
        }
        // ברירת מחדל – למשל "All Items"
        return 'All Items';
    };
    // 🆕 טעינת סדר שדות לתצוגה מתוך View של רשימת Integration
    var loadIntegrationViewFieldOrderForPhase = function (tenderPhaseRaw) { return __awaiter(void 0, void 0, void 0, function () {
        var viewName, vf, internalNames, e_2;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    console.log("🧩🧩🧩");
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, , 4]);
                    viewName = getIntegrationViewNameForTenderPhase(tenderPhaseRaw);
                    return [4 /*yield*/, sp.web.lists
                            .getById(INTEGRATION_LIST_ID)
                            .views.getByTitle(viewName)
                            .fields()];
                case 2:
                    vf = _a.sent();
                    internalNames = ((vf === null || vf === void 0 ? void 0 : vf.Items) || vf);
                    console.log('🧩 Integration view fields for', viewName, internalNames);
                    setIntegrationViewFieldOrder(internalNames || []);
                    return [3 /*break*/, 4];
                case 3:
                    e_2 = _a.sent();
                    console.error('Error loading Integration view fields for phase', e_2);
                    setIntegrationViewFieldOrder([]); // במקרה של שגיאה – נציג הכל
                    return [3 /*break*/, 4];
                case 4: return [2 /*return*/];
            }
        });
    }); };
    // מחזיר שם תצוגה (View) לפי הערך של TenderPhase ברשימת Integration
    var getViewNameForTenderPhase = function (rawPhase) {
        var phase = (rawPhase || ''); //.toLowerCase();
        if (phase.indexOf('Phase 1') === 0) {
            return 'Phase 1'; // שם ה-View ברשימת PMO
        }
        if (phase.indexOf('Phase 2') === 0) {
            return 'Phase 2'; // שם ה-View ברשימת PMO
        }
        if (phase.indexOf('Phase 3') === 0) {
            return 'Phase 3'; // שם ה-View ברשימת PMO
        }
        console.log("");
        // ברירת מחדל – למשל All Items או מה שתרצי
        return 'כל הפריטים';
    };
    // טוען את רשימת העמודות מה-View המתאים (לפי TenderPhase)
    var loadViewFieldOrderForPhase = function (tenderPhaseRaw) { return __awaiter(void 0, void 0, void 0, function () {
        var viewName, fields, resp, internalNames, e_3;
        var _a;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    console.log("🫏🐴🐎 loadViewFieldOrderForPhase");
                    _b.label = 1;
                case 1:
                    _b.trys.push([1, 3, , 4]);
                    viewName = getViewNameForTenderPhase(tenderPhaseRaw);
                    console.log("🫏🐴🐎 1");
                    setCurrentViewName(viewName);
                    console.log("🫏🐴🐎 2");
                    console.log("currentViewName ", currentViewName, " | ", viewName);
                    console.log("🫏🐴🐎 2.5");
                    return [4 /*yield*/, sp.web.lists
                            .getById(PMO_LIST_ID)
                            .views
                            .getById("62CFFAB1-507D-43B5-A87F-4F4CCD64BAD2")
                            .fields()];
                case 2:
                    fields = _b.sent();
                    console.log("🫏🐴🐎 3");
                    console.log("🫏🐴🐎 4");
                    resp = fields;
                    console.log("🫏🐴🐎 5");
                    internalNames = (resp.Items) || [];
                    console.log("🫏🐴🐎 6");
                    console.log("internalNames ", internalNames);
                    setViewFieldOrder(internalNames || []);
                    // אם את רוצה שכל טאב ב-Pivot יהיה בעצם "ה-View עצמו":
                    setStepsInternal((_a = {},
                        _a[viewName] = internalNames || [],
                        _a));
                    setActiveStep(viewName);
                    return [3 /*break*/, 4];
                case 3:
                    e_3 = _b.sent();
                    console.log("🫏🐴🐎 catch");
                    console.error('🫏🐴🐎 Error loading view fields for phase', e_3);
                    return [3 /*break*/, 4];
                case 4: return [2 /*return*/];
            }
        });
    }); };
    //🎀
    // האם להציג את RFCResponseLetterNo? רק אם RevisionIncludesChangeInTenderDo = 'Y'
    var showRFCResponseLetterNo = React.useMemo(function () {
        var _a;
        var v = String((_a = pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.RevisionIncludesChangeInTenderDo) !== null && _a !== void 0 ? _a : '').trim().toUpperCase();
        console.log("🎀[pmoDraft?.RevisionIncludesChangeInTenderDo]  -  v ", v);
        return v === 'YES';
    }, [pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.RevisionIncludesChangeInTenderDo]);
    React.useEffect(function () {
        console.log("🎍 in use effect trying to find your OLM");
        if (!sp || !(me === null || me === void 0 ? void 0 : me.Email))
            return;
        console.log("🎍 !sp || !me?.Email");
        var cancelled = false;
        (function () { return __awaiter(void 0, void 0, void 0, function () {
            var roles;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, myOLM(sp, "b321608b-6405-4ad0-8676-11c7350fa7a4", String(me.Email))];
                    case 1:
                        roles = _a.sent();
                        if (!cancelled) {
                            setMyRoles(roles);
                        }
                        return [2 /*return*/];
                }
            });
        }); })();
        return function () {
            cancelled = true;
        };
    }, [sp, me === null || me === void 0 ? void 0 : me.Email]);
    // אופציונלי: אם השדה מוסתר – ננקה את הערך כדי שלא יישמר בטעות
    React.useEffect(function () {
        if (!showRFCResponseLetterNo && (pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.RFCResponseLetterNo)) {
        }
    }, [showRFCResponseLetterNo, pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.RFCResponseLetterNo]);
    //🎀
    var loadFormForIntegration = function () { return __awaiter(void 0, void 0, void 0, function () {
        var integ, tenderPhaseRaw, _a, pmoFound, isNew, _b, permMap, general, tenderTeamUsers, pmoMaps, fieldMap, normalized, firstStep, tenderTeam;
        return __generator(this, function (_c) {
            switch (_c.label) {
                case 0:
                    if (!integrationId)
                        return [2 /*return*/];
                    return [4 /*yield*/, fetchIntegrationItemByGuid(sp, integrationId)];
                case 1:
                    integ = _c.sent();
                    setIntegrationItem(integ);
                    tenderPhaseRaw = String((integ === null || integ === void 0 ? void 0 : integ.TenderPhase) || '');
                    return [4 /*yield*/, loadViewFieldOrderForPhase(tenderPhaseRaw)];
                case 2:
                    _c.sent();
                    return [4 /*yield*/, loadIntegrationViewFieldOrderForPhase(tenderPhaseRaw)];
                case 3:
                    _c.sent();
                    // PMO לפי lookup ל-Integration
                    console.log("loadFormForIntegration", {
                        integrationId: integrationId,
                        typeofIntegrationId: typeof integrationId,
                        pmoIntegrationLookupName: pmoIntegrationLookupName,
                    });
                    return [4 /*yield*/, fetchOrCreatePmoByIntegration(sp, pmoListTitle, integrationId, pmoIntegrationLookupName)];
                case 4:
                    _a = _c.sent(), pmoFound = _a.item, isNew = _a.isNew;
                    setPmoItem(pmoFound);
                    setPmoDraft(pmoFound);
                    console.log("pmoFound ", pmoFound);
                    return [4 /*yield*/, Promise.all([
                            loadFieldPermissionMap(sp),
                            loadGeneralRoleUsers(sp),
                            null
                        ])];
                case 5:
                    _b = _c.sent(), permMap = _b[0], general = _b[1], tenderTeamUsers = _b[2];
                    setFieldPermMap(permMap);
                    setRoleUsers({
                        IntegrationTeam: general.IntegrationTeam,
                        IntegrationTeamLawyer: general.IntegrationTeamLawyer,
                        FinancialAdvisorIntegrationTeam: general.FinancialAdvisorIntegrationTeam,
                        M1TenderTeam: general.M1TenderTeam,
                        LawyerM1TenderTeam: general.LawyerM1TenderTeam,
                        M2TenderTeam: general.M2TenderTeam,
                        LawyerM2TenderTeam: general.LawyerM2TenderTeam,
                        M3TenderTeam: general.M3TenderTeam,
                        LawyerM3TenderTeam: general.LawyerM3TenderTeam,
                        FinancialAdvisor: general.FinancialAdvisor,
                        Lawyer: general.Lawyer,
                        PMOIntegrationTeam: general.PMOIntegrationTeam,
                        PMOTenderTeam: tenderTeamUsers ? tenderTeamUsers : [""], // ← חשוב: כאן נכנסת הרשאת "PMO צוות מכרז"
                    });
                    if (isNew)
                        setMsg({ type: MessageBarType.success, text: 'A new PMO item linked to Integration has been created.' });
                    else
                        setMsg(null);
                    return [4 /*yield*/, getFieldMapsById(sp, PMO_LIST_ID)];
                case 6:
                    pmoMaps = _c.sent();
                    setPmoLabels(pmoMaps.internalToTitle);
                    console.log("15🔮");
                    return [4 /*yield*/, getFieldInfoMapById(sp, PMO_LIST_ID)];
                case 7:
                    fieldMap = _c.sent();
                    console.log("16🔮 fieldMap ", fieldMap);
                    setPmoFieldInfoMap(fieldMap);
                    console.log("17🔮");
                    normalized = normalizeStepsConfigToInternal(stepsConfig, pmoMaps.titleToInternal);
                    console.log("18🔮");
                    setStepsInternal(normalized);
                    console.log("19🔮");
                    firstStep = Object.keys(normalized)[0] || '';
                    console.log("20🔮");
                    setActiveStep(firstStep);
                    console.log("21🔮");
                    return [4 /*yield*/, loadTenderTeamUsersFromIntegration(sp, integrationId)];
                case 8:
                    tenderTeam = _c.sent();
                    console.log("22🔮");
                    setRoleUsers(function (prev) { return (__assign(__assign({}, prev), { PMOTenderTeam: tenderTeam })); });
                    console.log("23🔮");
                    return [2 /*return*/];
            }
        });
    }); };
    useEffect(function () {
        (function () { return __awaiter(void 0, void 0, void 0, function () {
            var user, _a;
            return __generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, sp.web.currentUser.select('Email', 'Title', 'LoginName')()];
                    case 1:
                        user = _b.sent();
                        setMe(user);
                        return [3 /*break*/, 3];
                    case 2:
                        _a = _b.sent();
                        setMe(null);
                        return [3 /*break*/, 3];
                    case 3: return [2 /*return*/];
                }
            });
        }); })();
    }, [sp]);
    var AUTO_RELOAD_KEY = 'SmartForm_AutoReloadOnce';
    useEffect(function () {
        (function () { return __awaiter(void 0, void 0, void 0, function () {
            var e_4, didReload, guardErr_1;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        _a.trys.push([0, 4, 11, 12]);
                        setBusy(true);
                        // ✅ WAIT for SP to be fully initialized
                        return [4 /*yield*/, sp.web.select("Title")()];
                    case 1:
                        // ✅ WAIT for SP to be fully initialized
                        _a.sent(); // Simple call to ensure context is ready
                        return [4 /*yield*/, Promise.all([
                                loadIntegrationChoices(),
                                loadIntegrationMeta(),
                                loadIntegrationLabels(),
                                (function () { return __awaiter(void 0, void 0, void 0, function () {
                                    var fp, roles;
                                    return __generator(this, function (_a) {
                                        switch (_a.label) {
                                            case 0: return [4 /*yield*/, loadFieldPermissionMap(sp)];
                                            case 1:
                                                fp = _a.sent();
                                                setFieldPermMap(fp);
                                                return [4 /*yield*/, loadGeneralRoleUsers(sp)];
                                            case 2:
                                                roles = _a.sent();
                                                setRoleUsers(function (prev) { return (__assign(__assign({}, prev), { FinancialAdvisor: roles.FinancialAdvisor.map(function (e) { return e.toLowerCase(); }), Lawyer: roles.Lawyer.map(function (e) { return e.toLowerCase(); }), PMOIntegrationTeam: roles.PMOIntegrationTeam.map(function (e) { return e.toLowerCase(); }) })); });
                                                return [2 /*return*/];
                                        }
                                    });
                                }); })(),
                            ])];
                    case 2:
                        _a.sent();
                        sessionStorage.removeItem(AUTO_RELOAD_KEY);
                        return [4 /*yield*/, resetReloadGuard(sp)];
                    case 3:
                        _a.sent();
                        return [3 /*break*/, 12];
                    case 4:
                        e_4 = _a.sent();
                        _a.label = 5;
                    case 5:
                        _a.trys.push([5, 9, , 10]);
                        if (!!isEditMode) return [3 /*break*/, 8];
                        // ✅ Ensure SP is ready before reload guard operations
                        return [4 /*yield*/, sp.web.select("Title")()];
                    case 6:
                        // ✅ Ensure SP is ready before reload guard operations
                        _a.sent();
                        return [4 /*yield*/, markReloadedAndReload(sp)];
                    case 7:
                        didReload = _a.sent();
                        if (didReload)
                            return [2 /*return*/];
                        _a.label = 8;
                    case 8: return [3 /*break*/, 10];
                    case 9:
                        guardErr_1 = _a.sent();
                        console.log("Reload guard failed:", guardErr_1);
                        return [3 /*break*/, 10];
                    case 10:
                        setMsg({
                            type: MessageBarType.error,
                            text: 'Error loading initial data: ' + ((e_4 === null || e_4 === void 0 ? void 0 : e_4.message) || e_4) + "\n Please reload the page and everything will be ok."
                        });
                        return [3 /*break*/, 12];
                    case 11:
                        setBusy(false);
                        return [7 /*endfinally*/];
                    case 12: return [2 /*return*/];
                }
            });
        }); })();
    }, [sp]);
    useEffect(function () {
        (function () { return __awaiter(void 0, void 0, void 0, function () {
            var e_5, msgText, alreadyTried;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!integrationId)
                            return [2 /*return*/];
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, 4, 5]);
                        console.log("🥯1");
                        setBusy(true);
                        console.log("🥯2");
                        return [4 /*yield*/, loadFormForIntegration()];
                    case 2:
                        _a.sent();
                        console.log("🥯3");
                        return [3 /*break*/, 5];
                    case 3:
                        e_5 = _a.sent();
                        console.log("🥯4");
                        setMsg({ type: MessageBarType.error, text: 'Error loading the form: ' + ((e_5 === null || e_5 === void 0 ? void 0 : e_5.message) || e_5) });
                        msgText = String((e_5 === null || e_5 === void 0 ? void 0 : e_5.message) || e_5 || '');
                        setMsg({
                            type: MessageBarType.error,
                            text: 'Error loading initial data: ' + msgText
                        });
                        console.log("1");
                        // 🔁 אם זו בדיוק השגיאה של web, ננסה ריענון *פעם אחת בלבד*
                        if (msgText.indexOf("Cannot read properties of undefined (reading 'web')") != -1) {
                            alreadyTried = sessionStorage.getItem(AUTO_RELOAD_KEY);
                            console.log("1");
                            if (!alreadyTried) {
                                sessionStorage.setItem(AUTO_RELOAD_KEY, '1');
                                console.log("1");
                                window.location.reload();
                                console.log("1");
                            }
                        }
                        console.log("1");
                        return [3 /*break*/, 5];
                    case 4:
                        setBusy(false);
                        return [7 /*endfinally*/];
                    case 5:
                        console.log("1");
                        return [2 /*return*/];
                }
            });
        }); })();
    }, [integrationId]);
    var onChangeField = function (internal, value) {
        console.log("⛔ allowed");
        var toCsvString = function (v) {
            if (typeof v === "string")
                return v;
            if (Array.isArray(v))
                return v.join(",");
            if ((v === null || v === void 0 ? void 0 : v.results) && Array.isArray(v.results))
                return v.results.join(",");
            return String(v !== null && v !== void 0 ? v : "");
        };
        if (internal === 'DecisionAppliesToOtherWorksTende') {
            if (!allowedTenderTitlesReady) {
                console.warn("Tender titles not loaded yet - blocking change for now");
                return; // או תחליטי לא לחסום, אבל זה הכי בטוח
            }
            var vStr = toCsvString(value).trim();
            console.log("vStr 🦒 -", vStr, "-");
            //const vStr = String(value ?? '').trim();
            //console.log("vStr 🦒 ", vStr);
            var ok = isValidTenderSelection({
                input: vStr,
                validTitles: allowedTenderTitles,
            });
            console.log("ok ", ok);
            if (!ok) {
                console.warn("Blocked invalid value for DecisionAppliesToOtherWorksTende:", vStr);
                return; // ⛔ לא מעדכנים ל-draft
            }
            /*
            if (allowed.indexOf(vStr) === -1) {
              console.log("OI VEU 🦒😭🔮");
              console.warn('Blocked invalid value for DecisionAppliesToOtherWorksTende:', value);
        
              
              return; // ⛔ לא מעדכנים ל-draft
            }*/
        }
        // --- כאן האוטומציה החדשה ---
        setPmoDraft(function (prev) {
            var _a;
            var _b, _c;
            console.log("StatusOfRFCresponseOrTcRFC 🚌🚐🚎🚑🚒🚚🚛🚜🚘🚔🚖🚍🛻🚙🛺🚕🚓🚗");
            var next = __assign(__assign({}, (prev || {})), (_a = {}, _a[internal] = value, _a));
            var nowIso = new Date().toISOString();
            // 1) ActualDate מתעדכן כש-StatusOfRFCresponseOrTcRFC משתנה
            if (internal === 'StatusOfRFCresponseOrTcRFC') {
                console.log("🚌🚐🚎🚑🚒🚚🚛🚜🚘🚔🚖🚍🛻🚙🛺🚕🚓🚗");
                var prevVal = String((_b = prev === null || prev === void 0 ? void 0 : prev.StatusOfRFCresponseOrTcRFC) !== null && _b !== void 0 ? _b : '');
                var newVal = String(value !== null && value !== void 0 ? value : '');
                if (prevVal !== newVal) {
                    next.ActualDate = nowIso;
                }
            }
            // 2) DecisionDate מתעדכן כש-DecisionRegardingProposedChange משתנה
            if (internal === 'DecisionRegardingProposedChange') {
                var prevVal = String((_c = prev === null || prev === void 0 ? void 0 : prev.DecisionRegardingProposedChange) !== null && _c !== void 0 ? _c : '');
                var newVal = String(value !== null && value !== void 0 ? value : '');
                if (prevVal !== newVal) {
                    next.DecisionDate = nowIso;
                }
            }
            return next;
        });
        /*
        setPmoDraft((prev: any) => (
         
          { ...prev, [internal]: value }
        )
        );
      */
        // אם השדה הפך ללא ריק – ננקה שגיאת "שדה חובה"
        setValidationErrors(function (prev) {
            if (!prev[internal])
                return prev;
            var copy = __assign({}, prev);
            delete copy[internal];
            return copy;
        });
    };
    function isValidUrl(value) {
        try {
            new URL(value);
            return true;
        }
        catch (_a) {
            return false;
        }
    }
    var onSave = function (options) { return __awaiter(void 0, void 0, void 0, function () {
        var newErrors, flage, i, internal, v, isEmpty, info, t, urlToCheck, draftToSave, internal, info, t, val, urlStr, saved, syncErr_1, pmoList, integrationList, updatedPmoItem, e_6;
        var _a, _b;
        return __generator(this, function (_c) {
            switch (_c.label) {
                case 0:
                    console.log("🐴🐴🐴 in onsave pmoItem ", pmoItem);
                    if (!pmoItem || !pmoItem.Id)
                        return [2 /*return*/];
                    newErrors = {};
                    flage = false;
                    for (i = 0; i < REQUIRED_FIELDS.length; i++) {
                        internal = REQUIRED_FIELDS[i];
                        // אם אין למשתמש הרשאה לערוך את השדה – הוא לא שדה חובה עבורו
                        if (!canEditField(internal)) {
                            continue;
                        }
                        if (!isFieldVisibleNow(internal)) {
                            console.timeLog("🐴 !isFieldVisibleNow(internal) ", internal);
                            continue;
                        }
                        v = pmoDraft ? pmoDraft[internal] : undefined;
                        if (internal === "RFCresponseAsPublishedToBeFilled") {
                            console.log("RFCresponseAsPublishedToBeFilled v ", v);
                            if (v === false) {
                                flage = true;
                            }
                        }
                        if (flage === true) {
                            if (internal === "Addendum" || internal === "addendumDate") {
                                continue;
                            }
                            if (internal === "RFCResponseLetterNo") {
                                continue;
                            }
                            if (internal === "TenderCommitteeApprovalDate") {
                                console.log("🐴🐴 internal === TenderCommitteeApprovalDate and flage === true");
                                continue;
                            }
                        }
                        isEmpty = false;
                        if (v === null || v === undefined) {
                            console.log("🐴 internal ", internal);
                            isEmpty = true;
                        }
                        else if (typeof v === 'string') {
                            isEmpty = v.trim() === '';
                        }
                        else if (Array.isArray(v)) {
                            isEmpty = v.length === 0;
                        }
                        if (isEmpty) {
                            newErrors[internal] = 'Required field';
                        }
                        else if (!isEmpty) {
                            info = pmoFieldInfoMap[internal];
                            t = String((info === null || info === void 0 ? void 0 : info.TypeAsString) || '').toLowerCase();
                            if (t === 'url' || t === 'hyperlink') {
                                urlToCheck = "";
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
                        return [2 /*return*/]; // לא ממשיכים לשמירה ב־SharePoint
                    }
                    _c.label = 1;
                case 1:
                    _c.trys.push([1, 14, 15, 16]);
                    setBusy(true);
                    draftToSave = __assign({}, (pmoDraft || {}));
                    if (options === null || options === void 0 ? void 0 : options.updateEditingDate) {
                        draftToSave.IntegrationTeamDecisionEditingLa =
                            new Date().toISOString();
                    }
                    // המרה של שדה ההיפר־לינק לפורמט של SharePoint:
                    // { Url: 'https://...', Description: '...' }
                    // 💡 Normalize all URL / Hyperlink fields before saving
                    for (internal in pmoFieldInfoMap) {
                        if (!Object.prototype.hasOwnProperty.call(pmoFieldInfoMap, internal))
                            continue;
                        info = pmoFieldInfoMap[internal];
                        t = String((info === null || info === void 0 ? void 0 : info.TypeAsString) || '').toLowerCase();
                        // גם URL וגם Hyperlink
                        if (t !== 'url' && t !== 'hyperlink')
                            continue;
                        val = draftToSave[internal];
                        if (val === undefined || val === null) {
                            draftToSave[internal] = null;
                            continue;
                        }
                        // אם כבר בפורמט הנכון – לא נוגעים
                        if (typeof val === 'object' && (val.Url || val.url))
                            continue;
                        urlStr = String(val || '').trim();
                        if (!urlStr) {
                            draftToSave[internal] = null;
                        }
                        else {
                            draftToSave[internal] = {
                                Url: urlStr,
                                Description: urlStr,
                            };
                        }
                    }
                    console.log('🧾 draftToSave before save:', draftToSave);
                    return [4 /*yield*/, savePmoItem(sp, PMO_LIST_ID, pmoItem.Id, draftToSave)];
                case 2:
                    saved = _c.sent();
                    console.log("🍕 saved ", saved);
                    setPmoItem(saved);
                    setPmoDraft(saved);
                    setMsg({ type: MessageBarType.success, text: 'Saved successfully.' });
                    setValidationErrors({});
                    // ---- אם אין שגיאות חובה, ממשיכים לשמור ----
                    console.log("🦘saved ", saved);
                    _c.label = 3;
                case 3:
                    _c.trys.push([3, 7, , 8]);
                    draftToSave.DueDateCalculated = saved.DueDateCalculated;
                    if (!draftToSave.IntegrationId) return [3 /*break*/, 5];
                    //await syncPmoToIntegration(sp, integrationId, draftToSave);
                    console.log("🍕 draftToSave ", draftToSave);
                    return [4 /*yield*/, syncPmoToIntegration(sp, draftToSave.IntegrationId, draftToSave)];
                case 4:
                    _c.sent();
                    console.log('✅ Synced PMO → INTEGRATION for item', integrationId, "draftToSave \n", draftToSave);
                    return [3 /*break*/, 6];
                case 5:
                    console.warn('syncPmoToIntegration: integrationId is null – no sync done');
                    _c.label = 6;
                case 6: return [3 /*break*/, 8];
                case 7:
                    syncErr_1 = _c.sent();
                    console.error('❌ Failed to sync PMO → INTEGRATION', syncErr_1);
                    return [3 /*break*/, 8];
                case 8:
                    if (!((options === null || options === void 0 ? void 0 : options.updateEditingDate) === true)) return [3 /*break*/, 13];
                    console.log("updateEditingDate is true ");
                    pmoList = sp.web.lists.getById(PMO_LIST_ID);
                    return [4 /*yield*/, pmoList.items.getById(pmoItem.Id).update((_a = {},
                            _a[INTEGRATION_TEAM_STATUS_FIELD] = STATUS_DECISION_REACHED,
                            _a[PMO_SENT_PROTOCOL_FIELD] = true,
                            _a))];
                case 9:
                    _c.sent();
                    integrationList = sp.web.lists.getById(INTEGRATION_LIST_ID);
                    if (!integrationId) return [3 /*break*/, 11];
                    return [4 /*yield*/, integrationList.items.getById(integrationId).update((_b = {},
                            _b[INTEGRATION_TEAM_STATUS_FIELD] = STATUS_DECISION_REACHED,
                            _b))];
                case 10:
                    _c.sent();
                    _c.label = 11;
                case 11: return [4 /*yield*/, pmoList.items.getById(pmoItem.Id)()];
                case 12:
                    updatedPmoItem = _c.sent();
                    setPmoItem(updatedPmoItem);
                    setPmoDraft(updatedPmoItem);
                    setValidationErrors({});
                    // אופציה א: לשמור ל-state שמציג "כל הפריט"
                    //setLastSavedItem(updatedPmoItem);
                    // אופציה ב: אם את רוצה להציג אותו בתוך הטופס עצמו:
                    // setPmoDecisionItem(updatedPmoItem);
                    // אופציה ג: פשוט להדפיס לקונסול
                    console.log("✅ Updated PMO Decision item:", updatedPmoItem);
                    _c.label = 13;
                case 13: return [3 /*break*/, 16];
                case 14:
                    e_6 = _c.sent();
                    setMsg({ type: MessageBarType.error, text: 'Save failed: ' + ((e_6 === null || e_6 === void 0 ? void 0 : e_6.message) || e_6) });
                    return [3 /*break*/, 16];
                case 15:
                    setBusy(false);
                    return [7 /*endfinally*/];
                case 16: return [2 /*return*/];
            }
        });
    }); };
    var onSplitTenderClick = function () { return __awaiter(void 0, void 0, void 0, function () {
        var e_7;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, 3, 4]);
                    setIsSplitting(true);
                    return [4 /*yield*/, splitTenderAndCreateIntegrationItems({
                            sp: sp,
                            integrationListId: "2c962132-409d-4bf2-9440-3b3b6c7975a0",
                            workTendersListId: "a33ec5e6-86c0-439a-9eff-f5807b7764d9",
                            pmoItem: pmoDraft,
                            itegrationItem: integrationItem,
                            tenderSourceInternalName: "DecisionAppliesToOtherWorksTende",
                            pmoToIntegrationMap: PMO_TO_INTEGRATION_FIELD_MAP,
                            linkFieldInternalName: "NTA_x2019_s_x0020_reference",
                            linkValue: pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.IntegrationId,
                        })];
                case 1:
                    _a.sent();
                    if (integrationItem.coppiedFrom != null) {
                        setMsg === null || setMsg === void 0 ? void 0 : setMsg({ type: MessageBarType.error, text: "Not allowed to split splitted items." });
                    }
                    else {
                        setMsg === null || setMsg === void 0 ? void 0 : setMsg({ type: MessageBarType.success, text: "Split tender completed." });
                    }
                    return [3 /*break*/, 4];
                case 2:
                    e_7 = _a.sent();
                    console.error("splitTenderAndCreateIntegrationItems failed", e_7);
                    setMsg === null || setMsg === void 0 ? void 0 : setMsg({ type: MessageBarType.error, text: (e_7 === null || e_7 === void 0 ? void 0 : e_7.message) || "Split tender failed." });
                    return [3 /*break*/, 4];
                case 3:
                    setIsSplitting(false);
                    return [7 /*endfinally*/];
                case 4: return [2 /*return*/];
            }
        });
    }); };
    function formatDateDDMMYYYY(value) {
        if (!value)
            return '';
        var d = new Date(value);
        if (isNaN(d.getTime()))
            return String(value);
        var dd = ('0' + d.getDate()).slice(0 - 2);
        var mm = ('0' + (d.getMonth() + 1)).slice(0 - 2);
        var yyyy = d.getFullYear();
        return "".concat(dd, "/").concat(mm, "/").concat(yyyy);
    }
    var renderIntegrationReadonly = function () {
        if (!integrationItem)
            return null;
        // 🧩 אם יש View מוגדר – ניקח את הסדר ממנו; אחרת – כל השדות מהפריט
        var allKeys = Object.keys(integrationItem);
        // אם יש View – נשתמש בשדות מה-View; אחרת – כל השדות מהפריט
        var baseKeys = (integrationViewFieldOrder && integrationViewFieldOrder.length
            ? integrationViewFieldOrder //.filter(k => allKeys.indexOf(k)!= -1)
            : allKeys);
        console.log("integrationViewFieldOrder ", integrationViewFieldOrder);
        var keys = baseKeys.filter(function (k) {
            console.log(" - integration key ", k);
            if (isSystemField(k)) {
                console.log("🛹 is system");
                return false;
            }
            if (k === 'Id' || k === 'ID' || k === 'Title' || k === 'formCreator' || k === 'LM_x2019_sreference' || k === 'DocumentReference' || k === 'SectionName')
                return false; //k === 'DocumentName' ||
            var info = integrationFieldInfoMap[k];
            if (info && info.Hidden === true) {
                console.log("🙈k info.Hidden === true ");
                return false;
            }
            return true;
        });
        return (React.createElement(Stack, { tokens: { childrenGap: 10 } }, keys.map(function (k) {
            var raw = integrationItem[k];
            var info = integrationFieldInfoMap[k];
            var text;
            var typeStr = String((info === null || info === void 0 ? void 0 : info.TypeAsString) || '').toLowerCase();
            if (k === "RelatedBidderRFCsexist_x003f_" || k === "weretheRelatedBidderRFCsresponde") {
                text = '';
            }
            // 🎯 1. Boolean → Yes / No
            if (typeStr === 'boolean') {
                if (raw === true) {
                    text = 'Yes';
                }
                else if (raw === false) {
                    text = 'No';
                }
                else {
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
                text = String(raw !== null && raw !== void 0 ? raw : '');
            }
            console.log("🍋‍🟩integrationLabels[k]  ", integrationLabels[k]);
            return (React.createElement("div", { key: k, style: {
                    padding: 10,
                    background: 'rgba(15,23,42,0.02)',
                    borderRadius: 12,
                    border: '1px solid #e5e7eb',
                    boxShadow: '0 6px 18px rgba(15,23,42,0.06)'
                } },
                React.createElement("div", { style: { fontSize: 12, color: '#6b7280', marginBottom: 4 } }, integrationLabels[k] || k),
                React.createElement("div", { style: { fontWeight: 600, whiteSpace: 'pre-wrap', color: '#111827' } }, text)));
        })));
    };
    // --- לוגיקת תלות לשדה IntegrationTeamDecisionImplement ---
    var APPLIES_FIELD = 'DecisionAppliesToOtherWorksTende';
    var TARGET_FIELD = 'IntegrationTeamDecisionImplement';
    var DATE_FIELD = 'DateForIntegrationTeamDecisionIm';
    // האם "Applies" מאפשר עריכה?
    var appliesVal = String((pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft[APPLIES_FIELD]) || '').trim();
    var canEdit_ITDI = (appliesVal != '');
    var not_Relevant = (appliesVal === 'Not relevant to additional tenders');
    /* (appliesVal === 'All Infra 1 tenders' || appliesVal === 'Infra#1 DB - M3-WP2'|| appliesVal === 'Infra#1 DB - M2-WP3'
    || appliesVal === 'Infra#1 DB - M1-WP1 + WP2' || appliesVal ==='M3-WPO (Outer Boxes)' ||
    appliesVal.indexOf('All Infra 1 tenders') !== -1 ||
     appliesVal.indexOf('Infra#1 DB - M3-WP2') !== -1 || appliesVal.indexOf('Infra#1 DB - M2-WP3') !== -1 ||
      appliesVal.indexOf('Infra#1 DB - M1-WP1 + WP2') !== -1 || appliesVal.indexOf('M3-WPO (Outer Boxes)') !== -1);
  */
    console.log("🦧 canEdit_ITDI ", canEdit_ITDI, " appliesVal ", appliesVal);
    // הפקת תאריך לתצוגה
    var dateText = '';
    try {
        var rawDate = pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft[DATE_FIELD];
        if (rawDate) {
            var d = new Date(rawDate);
            if (!isNaN(d.getTime())) {
                dateText = d.toLocaleDateString('he-IL');
                console.log("🥨dateText ", dateText);
            }
        }
    }
    catch (_7) { }
    // לפני ה־itdiOptions, פעם אחת:
    var tenderPStr = String((integrationItem === null || integrationItem === void 0 ? void 0 : integrationItem.TenderPhase) || '').trim().toLowerCase();
    var isPhase2 = tenderPStr.indexOf('phase 2') !== -1;
    console.log("🥗canEdit_ITDI ", canEdit_ITDI, " pmoDraft?.['StatusOfRFCresponseOrTcRFC'] === 'Issued' ", (pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft['StatusOfRFCresponseOrTcRFC']) === 'Issued', " isPhase2 ", isPhase2, " ");
    console.log("🥨dateText ", dateText);
    var itdiOptions = not_Relevant
        ? [
            { key: 'Done', text: dateText ? "Done \u2014 ".concat(dateText) : 'Done' },
            { key: 'Pending', text: dateText ? "Pending \u2014 ".concat(dateText) : 'Pending' },
            { key: 'Not required', text: 'Not required' },
        ]
        : (canEdit_ITDI && ((isPhase2 && (pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft['StatusOfRFCresponseOrTcRFC']) === 'Issued') || !isPhase2))
            ? [
                { key: 'Done', text: dateText ? "Done \u2014 ".concat(dateText) : 'Done' },
                { key: 'Pending', text: dateText ? "Pending \u2014 ".concat(dateText) : 'Pending' },
            ]
            : canEdit_ITDI
                ? [
                    { key: 'Pending', text: dateText ? "Pending \u2014 ".concat(dateText) : 'Pending' },
                ]
                : [
                    { key: 'Not required', text: 'Not required' },
                ];
    /*const itdiOptions = (!not_Relevant && canEdit_ITDI && ((isPhase2 &&  pmoDraft?.['StatusOfRFCresponseOrTcRFC'] === 'Issued') ||(!isPhase2)))
    ? [
      { key: 'Done', text: dateText ? `Done — ${dateText}` : 'Done' },
      { key: 'Pending', text: dateText ? `Pending — ${dateText}` : 'Pending' }
    ]
    :canEdit_ITDI ? [
      { key: 'Pending', text: dateText ? `Pending — ${dateText}` : 'Pending' }
    ]
    : [
      { key: 'Not required', text: 'Not required' },
    ];*/
    useEffect(function () {
        var _a;
        var v = String((_a = pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.DecisionAppliesToOtherWorksTende) !== null && _a !== void 0 ? _a : '').trim();
        if (!v)
            return;
        var toCsvString = function (v) {
            if (typeof v === "string")
                return v;
            if (Array.isArray(v))
                return v.join(",");
            if ((v === null || v === void 0 ? void 0 : v.results) && Array.isArray(v.results))
                return v.results.join(",");
            return String(v !== null && v !== void 0 ? v : "");
        };
        var vStr = toCsvString(v).trim();
        console.log("vStr 🐻 ", vStr);
        if (!allowedTenderTitlesReady) {
            console.warn("Tender titles not loaded yet - blocking change for now");
            return; // או תחליטי לא לחסום, אבל זה הכי בטוח
        }
        var ok = isValidTenderSelection({
            input: vStr,
            validTitles: allowedTenderTitles,
        });
        if (!ok) {
            console.warn('Auto-fixing invalid DecisionAppliesToOtherWorksTende:', v);
            setPmoDraft(function (prev) { return (__assign(__assign({}, prev), { DecisionAppliesToOtherWorksTende: null })); });
            setMsg({
                type: MessageBarType.warning,
                text: "Field \"Decision Applies\" had invalid value \"".concat(v, "\" and was cleared.")
            });
        }
    }, [pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.DecisionAppliesToOtherWorksTende]);
    useEffect(function () {
        console.log("🍢 canEdit_ITDI ", canEdit_ITDI);
        if (!canEdit_ITDI) {
            console.log("🍥🎀🥩 canEdit_ITDI is not null - TARGET_FIELD", TARGET_FIELD, " pmoDraft?.[TARGET_FIELD] ", pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft[TARGET_FIELD]);
            if ((pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft[TARGET_FIELD]) !== 'Not required') {
                console.log("🪂");
                setPmoDraft(function (prev) {
                    var _a;
                    return (__assign(__assign({}, (prev || {})), (_a = {}, _a[TARGET_FIELD] = 'Not required', _a)));
                });
            }
        }
        console.log("in the end 🍢 canEdit_ITDI ", canEdit_ITDI);
    }, [canEdit_ITDI, pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft[TARGET_FIELD]]);
    var dynamicHideFieldstenders = React.useMemo(function () {
        var base = ['Integration', 'IntegrationId', 'Id', 'ID', 'Title'];
        console.log("🪺viewFieldOrder ", viewFieldOrder);
        // 1. יישור לפי ה-View: כל שדה שלא מופיע ב-View → מוסתר
        if (viewFieldOrder && viewFieldOrder.length && pmoDraft) {
            var allFields = Object.keys(pmoDraft);
            for (var i = 0; i < allFields.length; i++) {
                var f = allFields[i];
                /*
                if (f === 'RFCorTcRFCasPublishedByNTaNew') continue;*/
                if (base.indexOf(f) === -1 && viewFieldOrder.indexOf(f) === -1) {
                    base.push(f);
                }
            }
        }
        console.log("🏞️ base ", base);
        // 2. לוגיקה של RevisionIncludesChangeInTenderDo
        var rv = pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.RevisionIncludesChangeInTenderDo; //YES
        var isYes = (function () {
            if (typeof rv === 'boolean')
                return rv;
            var s = String(rv !== null && rv !== void 0 ? rv : '').trim().toLowerCase();
            return s === 'yes'; //true
        })();
        if (!isYes)
            base.push('RFCResponseLetterNo');
        console.log("🏞️ base ", base);
        // 3. 🔹 לוגיקה חדשה לפי sentProtocol
        var rawSent = pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.sentProtocol;
        var sentStr = String(rawSent !== null && rawSent !== void 0 ? rawSent : '').trim();
        // אם זה עמודת בחירה/טקסט עם "כן"
        var isSentYes = rawSent === true || // אם זה Yes/No (boolean)
            sentStr === 'כן' || // אם זה טקסט בעברית
            sentStr.toLowerCase() === 'yes'; // אם יבוא לך באנגלית בעתיד
        // אם זה *לא* "כן" → מסתירים Addendum ו-IntegrationTeamDecisionImplement
        if (!isSentYes) {
            if (base.indexOf('Addendum') === -1)
                base.push('Addendum');
            if (base.indexOf('addendumDate') === -1)
                base.push('addendumDate');
            console.log("PRTCOL WAS NOT SENT🏕️🏕️🏞️🏕️🏞️");
        }
        else {
            for (var i = base.length - 1; i >= 0; i--) {
                if (base[i] === 'Addendum')
                    base.splice(i, 1);
                if (base[i] === 'addendumDate')
                    base.splice(i, 1);
            }
            console.log("PRTCOL WAS SENT🏕️🏕️🏞️🏕️🏞️");
        }
        var tenderPhaseRaw = pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.TenderPhase;
        var tenderPhaseStr = String(tenderPhaseRaw !== null && tenderPhaseRaw !== void 0 ? tenderPhaseRaw : '').toLowerCase();
        // מספיק שאו שזה בדיוק "1" או שמופיע בו "phase 1"
        var isPhase1 = tenderPhaseStr === 'phase 1 - bidders’ requests for clarifications (rfcs) of tender documents' || tenderPhaseStr.indexOf('phase 1') != -1;
        if (isPhase1 && !base.includes('StatusOfRFCresponseOrTcRFC')) {
            base.push('StatusOfRFCresponseOrTcRFC');
        }
        return base;
    }, [pmoDraft, pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.RevisionIncludesChangeInTenderDo, pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.sentProtocol, viewFieldOrder]);
    // ✅ אילו שדות של Tender Team *כן* יוצגו (TENDER_TEAM_FIELDS פחות מה שמוסתר דינאמית)
    var tenderTeamVisibleFields = React.useMemo(function () {
        var dynHidden = dynamicHideFieldstenders || [];
        console.log("dynamicHideFieldstenders ", dynamicHideFieldstenders);
        console.log("TENDER_TEAM_FIELDS.filter(f => !dynHidden.includes(f)) ", TENDER_TEAM_FIELDS.filter(function (f) { return !dynHidden.includes(f); }));
        return TENDER_TEAM_FIELDS.filter(function (f) { return !dynHidden.includes(f); });
    }, [dynamicHideFieldstenders]);
    // ✅ hideFields עבור ה־EditableFields של Tender Team:
    // מסתיר *כל* השדות שלא נמצאות ב־tenderTeamVisibleFields
    var tenderTeamHideFields = React.useMemo(function () {
        var _a;
        var src = (_a = (pmoDraft || pmoItem)) !== null && _a !== void 0 ? _a : {};
        return Object.keys(src).filter(function (k) { return !tenderTeamVisibleFields.includes(k); });
    }, [pmoDraft, pmoItem, tenderTeamVisibleFields]);
    var dynamicHideFields = React.useMemo(function () {
        var base = ['Integration', 'IntegrationId', 'Id', 'ID', 'Title'];
        // 1. יישור לפי ה-View: כל שדה שלא מופיע ב-View → מוסתר
        if (viewFieldOrder && viewFieldOrder.length && pmoDraft) {
            var allFields = Object.keys(pmoDraft);
            for (var i = 0; i < allFields.length; i++) {
                var f = allFields[i];
                /*
                if (f === 'RFCorTcRFCasPublishedByNTaNew') continue;*/
                if (base.indexOf(f) === -1 && viewFieldOrder.indexOf(f) === -1) {
                    base.push(f);
                }
            }
        }
        for (var i = 0; i < TENDER_TEAM_FIELDS.length; i++) {
            var f = TENDER_TEAM_FIELDS[i];
            base.push(f);
        }
        // 2. לוגיקה של RevisionIncludesChangeInTenderDo
        var rv = pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.RevisionIncludesChangeInTenderDo;
        var isYes = (function () {
            if (typeof rv === 'boolean')
                return rv;
            var s = String(rv !== null && rv !== void 0 ? rv : '').trim().toLowerCase();
            return s === 'yes';
        })();
        if (!isYes)
            base.push('RFCResponseLetterNo');
        // 3. 🔹 לוגיקה חדשה לפי sentProtocol
        var rawSent = pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.sentProtocol;
        var sentStr = String(rawSent !== null && rawSent !== void 0 ? rawSent : '').trim();
        // אם זה עמודת בחירה/טקסט עם "כן"
        var isSentYes = rawSent === true || // אם זה Yes/No (boolean)
            sentStr === 'כן' || // אם זה טקסט בעברית
            sentStr.toLowerCase() === 'yes'; // אם יבוא לך באנגלית בעתיד
        // אם זה *לא* "כן" → מסתירים Addendum ו-IntegrationTeamDecisionImplement
        if (!isSentYes) {
            console.log("protole wasn't sent yet ");
            if (base.indexOf('Addendum') === -1)
                base.push('Addendum');
            if (base.indexOf('addendumDate') === -1)
                base.push('addendumDate');
            //if (base.indexOf(TARGET_FIELD) === -1) base.push(TARGET_FIELD);
        }
        else {
            console.log("protole was sent ");
            for (var i = base.length - 1; i >= 0; i--) {
                if (base[i] === 'Addendum')
                    base.splice(i, 1);
                if (base[i] === 'addendumDate')
                    base.splice(i, 1);
            }
            console.log("base ", base);
        }
        return base;
    }, [pmoDraft, pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.RevisionIncludesChangeInTenderDo, pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.sentProtocol, viewFieldOrder]);
    var isFieldVisibleNow = function (internal) {
        var _a;
        if (['Integration', 'IntegrationId', 'Id', 'ID', 'Title'].indexOf(internal) > -1) {
            return false;
        }
        if (dynamicHideFields.indexOf(internal) > -1) {
            console.log("⛑️ dynamicHideFields  -  internal ", internal);
            return false;
        }
        var sentVal = String((_a = pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.sentProtocol) !== null && _a !== void 0 ? _a : '').trim();
        var isSentYes = (sentVal === 'true');
        console.log("💛sentVal ", sentVal);
        // אם זה אחד השדות הרלוונטיים, והוא לא "כן" → להסתיר
        if ((internal === 'Addendum' || internal === 'addendumDate') && !isSentYes) {
            console.log("🏕️🏕️🏞️🏕️🏞️");
            return false;
        }
        var decision = String((pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.DecisionRegardingProposedChange) || '').trim();
        var revInc = String((pmoDraft === null || pmoDraft === void 0 ? void 0 : pmoDraft.RevisionIncludesChangeInTenderDo) || '').trim().toLowerCase();
        console.log("🏕️ revInc ", revInc);
        var tenderPhaseStr = String((integrationItem === null || integrationItem === void 0 ? void 0 : integrationItem.TenderPhase) || '').trim().toLowerCase();
        // האם RevisionIncludesChangeInTenderDo הוא "כן"/True
        /*const revIsTrue = (() => {
          const rv = pmoDraft?.RevisionIncludesChangeInTenderDo;
          if (typeof rv === 'boolean') return rv;
          const s = String(rv ?? '').trim().toLowerCase();
          return  s === 'yes';
        })();*/
        if (internal === 'RFCresponseAsPublishedToBeFilled' || internal === 'StatusOfRFCresponseOrTcRFC') {
            if (decision !== 'Accept' && decision !== 'Partially accepted')
                return false;
        }
        if (internal === 'Addendum' || internal === 'addendumDate' || internal === 'TenderCommitteeApprovalDate') {
            console.log("🤩 1 internal ", internal);
            var isNo = (revInc.toLowerCase() != 'yes');
            if (isNo && isPhase2) {
                console.log("🐴🐴🐴🐴🐴 internal === 'TenderCommitteeApprovalDate' and isNo false");
                return false;
            }
            /*if (!revIsTrue || !isPhase2) {
              console.log("🐴🐴🐴🐴  isFieldVisibleNow TenderCommitteeApprovalDate");
              return false;
            }*/
        }
        if (internal === 'dog') {
            return false;
        }
        if (internal === 'RevisedWordingFinalForPublicatio') {
            if (tenderPhaseStr !== 'phase 1 – preparation of tender documents')
                return false;
            //return false;
        }
        if (internal === 'RevisionIncludesChangeInTenderDo' ||
            internal === 'RFCResponseLetterNo' ||
            internal === 'RFCresponseAsPublishedToBeFilled') {
            if (tenderPhaseStr !== 'phase 2 - bidders’ requests for clarifications (rfcs) of tender documents') {
                return false;
            }
        }
        if (internal === "formCreator") {
            return false;
        }
        return true;
    };
    var isFieldInTender = function (internal) {
        if (TENDER_TEAM_FIELDS.indexOf(internal) > -1) {
            return true;
        }
        return false;
    };
    var canEditField = function (internal) {
        console.log("🧨internal ", internal);
        var emailLc = String((me === null || me === void 0 ? void 0 : me.Email) || '').toLowerCase();
        if (!emailLc)
            return false;
        var olm = getOriginatingLineManager();
        console.log("🧨internal ", internal);
        if (internal === TARGET_FIELD) {
            console.log("🧨🎇 canEdit_ITDI ", canEdit_ITDI);
            /*return canEdit_ITDI && canUserEditField(emailLc, internal, fieldPermMap, {
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
          );*/
            return canUserEditField(emailLc, internal, fieldPermMap, {
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
            }, olm);
        }
        return canUserEditField(emailLc, internal, fieldPermMap, roleUsers, olm);
    };
    var renderPmoEditableBySteps = function () {
        var _a;
        // 🐛 DEBUG – print all PMO fields & visibilityte
        if (pmoDraft) {
            var allFields = Object.keys(pmoDraft || {});
            console.log('PMO all fields (internal names):', allFields);
            console.log('PMO viewFieldOrder:', viewFieldOrder);
            console.log('PMO dynamicHideFields:', dynamicHideFields);
            var temp = allFields.filter(function (f) { return isFieldVisibleNow(f); });
            console.log("🦁temp ", temp);
            var visibleNow = temp.filter(function (f) { return !isFieldInTender(f); });
            console.log("🔮 visibleNow ", visibleNow);
            console.log('PMO fields that isFieldVisibleNow() == true:', visibleNow);
        }
        // 🔹 נחלץ את המפתח של ה־Phase מתוך TenderPhase
        var phaseKey = getPhaseViewKeyFromTenderPhase(integrationItem === null || integrationItem === void 0 ? void 0 : integrationItem.TenderPhase);
        var allSteps = Object.keys(stepsInternal);
        // 🔹 אם יש Phase מזוהה – נסנן רק את ה־steps שהשם שלהם מתחיל בו
        var stepNames = phaseKey
            ? allSteps.filter(function (s) { return s.toLowerCase().indexOf(phaseKey) === 0; })
            : allSteps;
        // אם אין בכלל steps – מציגים טופס מלא בלי Pivot
        if (!stepNames.length) {
            var rawOrder = viewFieldOrder && viewFieldOrder.length
                ? viewFieldOrder
                : Object.keys(pmoDraft || {});
            var order = rawOrder.filter(function (k) { return (k !== "Addendum" && k !== "addendumDate"); });
            console.log("🧤1");
            return (React.createElement(EditableFields, { item: pmoDraft, onChange: onChangeField, fieldOrder: order, hideFields: dynamicHideFields, internalToTitle: pmoLabels, fieldInfoMap: pmoFieldInfoMap, canEdit: canEditField, placeholderMap: placeholders, choiceOverrides: (_a = {},
                    _a[TARGET_FIELD] = itdiOptions,
                    _a), tenderPhase: String((integrationItem === null || integrationItem === void 0 ? void 0 : integrationItem.TenderPhase) || ''), requiredMap: requiredMap, errorMap: validationErrors, labelOverrides: PMO_LABEL_OVERRIDES }));
        }
        var getStepFieldOrder = function (step) {
            var stepFields = stepsInternal[step] || [];
            // אם יש View – נכבד אותו קודם, ואז נסנן לפי ה-step
            var order;
            if (viewFieldOrder && viewFieldOrder.length) {
                if (stepFields.length) {
                    order = viewFieldOrder.filter(function (f) { return stepFields.indexOf(f) !== -1; });
                }
                else {
                    order = viewFieldOrder;
                }
            }
            else {
                // בלי View – נשאר רק עם ה-step המקורי
                order = stepFields;
            }
            return order; //mapped;
        };
        // לוודא שה־selectedKey תמיד שייך ל־stepNames
        var effectiveActiveStep = stepNames.indexOf(activeStep) != -1
            ? activeStep
            : (stepNames[0] || '');
        console.log("🧤2");
        return (React.createElement(Pivot, { selectedKey: effectiveActiveStep, onLinkClick: function (i) { return setActiveStep((i === null || i === void 0 ? void 0 : i.props.itemKey) || ''); }, styles: {
                root: {
                    marginTop: 8
                },
                link: {
                    fontWeight: 600,
                    fontSize: 14
                }
            } }, stepNames.map(function (step) {
            var _a;
            return (React.createElement(PivotItem, { headerText: step, itemKey: step, key: step },
                React.createElement("div", { style: { marginTop: 10 } },
                    React.createElement(EditableFields, { item: pmoDraft, onChange: onChangeField, fieldOrder: getStepFieldOrder(step), hideFields: dynamicHideFields, internalToTitle: pmoLabels, fieldInfoMap: pmoFieldInfoMap, canEdit: canEditField, placeholderMap: placeholders, choiceOverrides: (_a = {},
                            _a[TARGET_FIELD] = itdiOptions,
                            _a), tenderPhase: String((integrationItem === null || integrationItem === void 0 ? void 0 : integrationItem.TenderPhase) || ''), requiredMap: requiredMap, errorMap: validationErrors, labelOverrides: PMO_LABEL_OVERRIDES }))));
        })));
    };
    //integrationChoices
    //fieldPermMap, pmoDraft, pmoItem, integrationItem
    console.log("🍋‍🟩 RFCResponseLetterNo", TENDER_TEAM_FIELDS.indexOf('RFCResponseLetterNo') > -1 && tenderTeamHideFields.indexOf('RFCResponseLetterNo') <= -1);
    console.log("🍋 Addendum", TENDER_TEAM_FIELDS.indexOf('Addendum') > -1 && tenderTeamHideFields.indexOf('Addendum') <= -1);
    console.log("🍓 TenderCommitteeApprovalDate", TENDER_TEAM_FIELDS.indexOf('TenderCommitteeApprovalDate') > -1 && tenderTeamHideFields.indexOf('TenderCommitteeApprovalDate') <= -1);
    console.log("🍇 StatusOfRFCresponseOrTcRFC", TENDER_TEAM_FIELDS.indexOf('StatusOfRFCresponseOrTcRFC') > -1 && tenderTeamHideFields.indexOf('StatusOfRFCresponseOrTcRFC') <= -1);
    console.log("🍓🍋🍇🍋‍🟩tenderTeamHideFields", tenderTeamHideFields);
    // למעלה בקומפוננטה (לפני ה-return)
    return (React.createElement("div", { dir: "ltr", style: {
            background: PAGE_BG,
            minHeight: '100vh',
            padding: '24px 32px',
            textAlign: 'left'
        } },
        React.createElement(Stack, { tokens: { childrenGap: 20 }, styles: {
                root: {
                    maxWidth: 1400,
                    margin: '0 auto'
                }
            } },
            React.createElement(Stack, { horizontal: true, horizontalAlign: "space-between", verticalAlign: "center", tokens: { childrenGap: 12 } },
                React.createElement(Stack, { tokens: { childrenGap: 4 } },
                    React.createElement("h1", { style: { margin: 0, fontSize: 28, color: '#0f172a' } }, "Decision Form"))),
            msg ? (React.createElement(MessageBar, { messageBarType: msg.type, styles: { root: { borderRadius: 10, boxShadow: '0 10px 30px rgba(15,23,42,0.08)' } } }, msg.text)) : null,
            React.createElement("div", { style: {
                    background: CARD_BG,
                    borderRadius: CARD_RADIUS,
                    boxShadow: CARD_SHADOW,
                    padding: 18,
                    border: '1px solid rgba(148,163,184,0.25)'
                } },
                React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 12 }, style: { alignItems: 'flex-end' } },
                    React.createElement(Stack, { grow: true, tokens: { childrenGap: 8 } },
                        React.createElement("span", { style: {
                                fontSize: 13,
                                color: '#64748b'
                            } }),
                        React.createElement("div", { style: { fontSize: 18, fontWeight: 600, color: '#0f172a' } }, "Tender item selection"),
                        React.createElement(Dropdown, { label: "Filter by Originating Line Manager", selectedKey: olmFilter, onChange: function (_, opt) { return setOlmFilter((opt === null || opt === void 0 ? void 0 : opt.key) || 'ALL'); }, options: [
                                { key: 'ALL', text: 'All line managers' },
                                { key: 'M1', text: 'M1' },
                                { key: 'M2', text: 'M2' },
                                { key: 'M3', text: 'M3' },
                            ], styles: { root: { maxWidth: 220 } } }),
                        React.createElement(ComboBox, { label: "Search Integration by NTA reference", placeholder: busy ? "Loading…" : "Start typing to search…", options: filteredIntegrationChoices, selectedKey: integrationId !== null && integrationId !== void 0 ? integrationId : null, text: integrationSearch, autoComplete: "off", allowFreeform: true, openOnKeyboardFocus: true, useComboBoxAsMenuWidth: true, 
                            // לא לאפס כאן! זה שובר בחירה מרשימה.
                            // onFocus / onClick הוסרו בכוונה.
                            onInputValueChange: function (text) {
                                var _a;
                                var t = text !== null && text !== void 0 ? text : "";
                                setIntegrationSearch(t);
                                // אם המשתמש מתחיל להקליד משהו שלא שווה בדיוק לטקסט של הבחירה הקודמת → מנקים בחירה
                                if (integrationId != null) {
                                    var chosen = integrationChoices.find(function (o) { return o.key === integrationId; });
                                    var chosenText = String((_a = chosen === null || chosen === void 0 ? void 0 : chosen.text) !== null && _a !== void 0 ? _a : "");
                                    if (chosenText !== t)
                                        setIntegrationId(null);
                                }
                                // אופציונלי (מומלץ): אם יש התאמה מלאה לטקסט של אופציה – נבחר אותה אוטומטית
                                /*const tt = t.trim().toLowerCase();
                                if (!tt) return;
                                const exact = integrationChoices.find(
                                  o => String(o.text ?? "").trim().toLowerCase() === tt
                                );
                                if (exact) setIntegrationId(exact.key as number);*/
                            }, onChange: function (_, opt, __, value) {
                                var _a;
                                if (opt) {
                                    // בחירה מתוך הרשימה (קליק/אנטר)
                                    setIntegrationId(opt.key);
                                    setIntegrationSearch(String((_a = opt.text) !== null && _a !== void 0 ? _a : ""));
                                    return;
                                }
                                // Freeform: המשתמש רק הקליד טקסט
                                var t = value !== null && value !== void 0 ? value : "";
                                setIntegrationSearch(t);
                                // אם זה בדיוק שם של אופציה – נשמור ID, אחרת נשאיר null
                                var tt = t.trim().toLowerCase();
                                var exact = integrationChoices.find(function (o) { var _a; return String((_a = o.text) !== null && _a !== void 0 ? _a : "").trim().toLowerCase() === tt; });
                                setIntegrationId(exact ? exact.key : null);
                            }, styles: {
                                root: {
                                    width: "100%",
                                    maxWidth: 540,
                                    minWidth: 0,
                                },
                                label: { fontWeight: 600 },
                            } })),
                    React.createElement(DefaultButton, { text: "Refresh list", onClick: loadIntegrationChoices, disabled: busy, styles: {
                            root: { borderRadius: 999, paddingInline: 18 },
                        } }))),
            React.createElement(Stack, { horizontal: true, wrap: true, tokens: { childrenGap: 24 }, styles: { root: { alignItems: 'flex-start' } } },
                React.createElement(Stack, { grow: true, styles: {
                        root: {
                            background: CARD_BG,
                            borderRadius: CARD_RADIUS,
                            boxShadow: CARD_SHADOW,
                            padding: 18,
                            border: '1px solid rgba(148,163,184,0.25)',
                            minWidth: 0
                        }
                    }, tokens: { childrenGap: 10 } },
                    React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'baseline' } },
                        React.createElement("h3", { style: { margin: 0, fontSize: 18, color: ACCENT } }, "Integration \u2013 Display"),
                        React.createElement("span", { style: { fontSize: 12, color: '#9ca3af' } }, "Read only")),
                    React.createElement("p", { style: { margin: '0 0 8px', fontSize: 13, color: '#6b7280' } }),
                    renderIntegrationReadonly()),
                React.createElement(Stack, { grow: true, styles: {
                        root: {
                            background: CARD_BG,
                            borderRadius: CARD_RADIUS,
                            boxShadow: CARD_SHADOW,
                            padding: 18,
                            border: '1px solid rgba(148,163,184,0.25)',
                            minWidth: 0
                        }
                    }, tokens: { childrenGap: 10 } },
                    React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'baseline' } },
                        React.createElement("h3", { style: { margin: 0, fontSize: 18, color: ACCENT } }, "Integration Team Decision \u2013 Editing"),
                        React.createElement("span", { style: { fontSize: 12, color: '#9ca3af' } }, "Required fields are marked in red \u2022 Permissions by role")),
                    renderPmoEditableBySteps(),
                    React.createElement("div", { style: { marginTop: 16, display: 'flex', justifyContent: 'flex-end' } },
                        React.createElement("div", { style: { display: "flex", gap: 8, alignItems: "center" } },
                            React.createElement(PrimaryButton, { text: busy ? 'Saving…' : 'Save Integration Decision', disabled: busy || !pmoItem, onClick: function () { return onSave({ updateEditingDate: true }); }, styles: {
                                    root: {
                                        borderRadius: 999,
                                        paddingInline: 24,
                                        fontWeight: 600,
                                    },
                                } }),
                            React.createElement(DefaultButton, { text: isSplitting ? "Splitting..." : "Split tender", onClick: onSplitTenderClick, disabled: isSplitting || !integrationItem }))))),
            React.createElement(Stack, { styles: {
                    root: {
                        background: CARD_BG,
                        borderRadius: CARD_RADIUS,
                        boxShadow: CARD_SHADOW,
                        padding: 18,
                        border: '1px solid rgba(148,163,184,0.25)',
                        marginTop: 16
                    }
                }, tokens: { childrenGap: 10 } },
                React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'baseline' } },
                    React.createElement("h3", { style: { margin: 0, fontSize: 18, color: ACCENT } }, "Tender Team status"),
                    React.createElement("span", { style: { fontSize: 12, color: '#9ca3af' } }, "These fields are managed by the Tender Team and saved to the same PMO item.")),
                React.createElement(EditableFields, { item: pmoDraft || pmoItem || {}, 
                    /*onChange={(internalName: string, value: any) => {
                      setPmoDraft((prev: any) => ({
                        ...(prev || pmoItem || {}),
                        [internalName]: value,
                      }));
                    }}*/
                    onChange: function (internalName, value) {
                        setPmoDraft(function (prev) {
                            var _a;
                            var _b, _c;
                            console.log("INSIDE setPmoDraft updater"); // בדיקה 2
                            var base = prev || pmoItem || {};
                            var next = __assign(__assign({}, base), (_a = {}, _a[internalName] = value, _a));
                            var nowIso = new Date().toISOString();
                            if (internalName === "StatusOfRFCresponseOrTcRFC") {
                                var prevVal = String((_b = base === null || base === void 0 ? void 0 : base.StatusOfRFCresponseOrTcRFC) !== null && _b !== void 0 ? _b : "");
                                var newVal = String(value !== null && value !== void 0 ? value : "");
                                if (prevVal !== newVal)
                                    next.ActualDate = nowIso;
                            }
                            if (internalName === "DecisionRegardingProposedChange") {
                                var prevVal = String((_c = base === null || base === void 0 ? void 0 : base.DecisionRegardingProposedChange) !== null && _c !== void 0 ? _c : "");
                                var newVal = String(value !== null && value !== void 0 ? value : "");
                                if (prevVal !== newVal)
                                    next.DecisionDate = nowIso;
                            }
                            return next;
                        });
                    }, fieldInfoMap: pmoFieldInfoMap, fieldOrder: TENDER_TEAM_FIELDS, hideFields: tenderTeamHideFields, labelOverrides: pmoLabels, canEdit: canEditField, tenderPhase: String((integrationItem === null || integrationItem === void 0 ? void 0 : integrationItem.TenderPhase) || ''), choiceOverrides: (_b = {},
                        _b[TARGET_FIELD] = itdiOptions,
                        _b), placeholderMap: placeholders })),
            React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 10 }, horizontalAlign: "space-between", verticalAlign: "center" },
                React.createElement("span", { style: { fontSize: 12, color: '#6b7280' } },
                    "Form changes are saved directly to SharePoint when you click",
                    React.createElement("strong", null, "Save"),
                    "."),
                React.createElement(Stack, { horizontal: true, tokens: { childrenGap: 8 }, verticalAlign: "center" },
                    React.createElement(PrimaryButton, { text: busy ? 'Saveing' : 'Save', onClick: function () { return onSave({ updateEditingDate: false }); }, disabled: busy || !pmoItem, styles: {
                            root: { borderRadius: 999, paddingInline: 26, fontWeight: 600 }
                        } }),
                    React.createElement(DefaultButton, { text: "Refresh Form", onClick: loadFormForIntegration, disabled: busy || !integrationId, styles: {
                            root: { borderRadius: 999, paddingInline: 18 }
                        } }),
                    msg && (React.createElement("span", { style: {
                            marginLeft: 12,
                            padding: '4px 10px',
                            borderRadius: 999,
                            fontSize: 13,
                            fontWeight: 600,
                            backgroundColor: msg.type === MessageBarType.success ? '#ecfdf3' : '#fef2f2',
                            color: msg.type === MessageBarType.success ? '#166534' : '#b91c1c',
                            border: "1px solid ".concat(msg.type === MessageBarType.success ? '#bbf7d0' : '#fecaca'),
                            whiteSpace: 'nowrap',
                        } }, msg.text)))))));
};
export default FormApp;
//# sourceMappingURL=FormApp.js.map