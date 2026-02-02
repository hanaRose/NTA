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
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
//import { SPHttpClient } from '@microsoft/sp-http';
import { INTEGRATION_LIST_ID, EXCLUDED_INTEGRATION_FIELDS } from './internalNames';
import { isSystemField } from './constants';
import { spPost } from "@pnp/sp";
var lc = function (s) { return String(s || '').toLowerCase(); };
//const strEq = (a?: string, b?: string) => lc(a) === lc(b);
//const strHas = (a?: string, frag?: string) => lc(a).indexOf(lc(frag)) > -1;
export var FIELD_PERMISSION_LIST_TITLE = 'fieldPermission';
export var FIELD_PERMISSION_LIST_ID = '05b813e8-e560-476d-89a7-5eac957bdc38';
export var FP_COL_INTERNAL = 'internalFieldName'; // Text
export var FP_COL_WHO_CAN_EDIT = 'WhoCanEdite'; // Choice/MultiChoice (טקסטים בעברית לפי התיאור)
export var GENERAL_ROLE_DEF_LIST_TITLE = 'GeneralRoleDefinition';
export var COL_FINANCIAL_ADVISOR = 'FinancialAdvisor'; // People (multi)
export var COL_LAWYER = 'Lawyer'; // People (multi)
export var COL_PMO_INTEGRATION_TEAM = 'PMOIntegrationTeam'; // People (multi)
export var COL_PMO_TENDER_TEAM_FROM_CREATER = 'formCreator'; // People (multi) ברשימת Integration
function integrationList(sp) {
    return sp.web.lists.getById(INTEGRATION_LIST_ID);
}
export function getFieldMapsByTitle(sp, listTitle) {
    return __awaiter(this, void 0, void 0, function () {
        var fields, internalToTitle, titleToInternal, i, f, disp;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, sp.web.lists.getByTitle(listTitle).fields
                        .select('Title', 'InternalName')()];
                case 1:
                    fields = _a.sent();
                    internalToTitle = {};
                    titleToInternal = {};
                    for (i = 0; i < fields.length; i++) {
                        f = fields[i];
                        if (!f)
                            continue;
                        disp = f.Title || f.InternalName;
                        internalToTitle[f.InternalName] = disp;
                        titleToInternal[disp] = f.InternalName;
                    }
                    return [2 /*return*/, { internalToTitle: internalToTitle, titleToInternal: titleToInternal }];
            }
        });
    });
}
export function getFieldInfoMapById(sp, listId) {
    return __awaiter(this, void 0, void 0, function () {
        var fields, map, _i, fields_1, f;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, sp.web.lists
                        .getById(listId)
                        .fields
                        .select('InternalName', 'Title', 'TypeAsString', 'ReadOnlyField', 'Hidden', 'RichText', 'Choices', 'AllowMultipleValues')()];
                case 1:
                    fields = _a.sent();
                    map = {};
                    for (_i = 0, fields_1 = fields; _i < fields_1.length; _i++) {
                        f = fields_1[_i];
                        if (f === null || f === void 0 ? void 0 : f.InternalName) {
                            map[f.InternalName] = f;
                        }
                    }
                    return [2 /*return*/, map];
            }
        });
    });
}
export function getFieldMapsById(sp, listId) {
    return __awaiter(this, void 0, void 0, function () {
        var fields, internalToTitle, titleToInternal, i, f, disp;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, sp.web.lists.getById(listId).fields
                        .select('Title', 'InternalName')()];
                case 1:
                    fields = _a.sent();
                    internalToTitle = {};
                    titleToInternal = {};
                    for (i = 0; i < fields.length; i++) {
                        f = fields[i];
                        if (!f)
                            continue;
                        disp = f.Title || f.InternalName;
                        internalToTitle[f.InternalName] = disp;
                        titleToInternal[disp] = f.InternalName;
                    }
                    return [2 /*return*/, { internalToTitle: internalToTitle, titleToInternal: titleToInternal }];
            }
        });
    });
}
export function fetchIntegrationItemByGuid(sp, integrationItemId) {
    return __awaiter(this, void 0, void 0, function () {
        var list, item, arr, cleaned, keys, i, k;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    console.log("🔮a");
                    list = integrationList(sp);
                    console.log("🔮b");
                    if (!integrationItemId) return [3 /*break*/, 2];
                    console.log("🔮c");
                    return [4 /*yield*/, list.items.getById(integrationItemId).select('*')()];
                case 1:
                    item = _a.sent();
                    return [3 /*break*/, 4];
                case 2:
                    console.log("🔮d");
                    return [4 /*yield*/, list.items.orderBy('Id', false).top(1)()];
                case 3:
                    arr = _a.sent();
                    console.log("🔮e");
                    item = arr && arr.length ? arr[0] : null;
                    console.log("🔮f");
                    _a.label = 4;
                case 4:
                    if (!item)
                        return [2 /*return*/, null];
                    console.log("🔮g");
                    cleaned = {};
                    console.log("🔮h");
                    keys = Object.keys(item);
                    console.log("🔮i");
                    for (i = 0; i < keys.length; i++) {
                        k = keys[i];
                        if (isSystemField(k))
                            continue;
                        if (EXCLUDED_INTEGRATION_FIELDS.has(k))
                            continue;
                        cleaned[k] = item[k];
                    }
                    console.log("🔮k");
                    return [2 /*return*/, cleaned];
            }
        });
    });
}
export function resolveInternalName(sp, listTitle, displayOrInternal) {
    return __awaiter(this, void 0, void 0, function () {
        var fields, internal, needleLc, i, f, t, n;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, sp.web.lists.getByTitle(listTitle).fields
                        .select('Title', 'InternalName')()];
                case 1:
                    fields = _a.sent();
                    internal = displayOrInternal;
                    needleLc = lc(displayOrInternal);
                    for (i = 0; i < fields.length; i++) {
                        f = fields[i];
                        if (!f)
                            continue;
                        t = lc(f.Title || '');
                        n = lc(f.InternalName || '');
                        if (n === needleLc || t === needleLc) {
                            internal = f.InternalName;
                            break;
                        }
                    }
                    return [2 /*return*/, internal];
            }
        });
    });
}
export function getFieldInfoMap(sp, listTitle) {
    return __awaiter(this, void 0, void 0, function () {
        var fields, map, i, f;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, sp.web.lists.getByTitle(listTitle).fields
                        .select('InternalName', 'Title', 'TypeAsString', 'ReadOnlyField', 'Hidden', 'RichText', 'Choices', 'AllowMultipleValues')()];
                case 1:
                    fields = _a.sent();
                    map = {};
                    for (i = 0; i < fields.length; i++) {
                        f = fields[i];
                        if (f && f.InternalName)
                            map[f.InternalName] = f;
                    }
                    return [2 /*return*/, map];
            }
        });
    });
}
function cleanRichText(html) {
    if (!html)
        return null; // במקום להחזיר undefined — מחזירים null תמיד
    var cleaned = html;
    cleaned = cleaned.replace(/class="ExternalClass[^"]*"/g, "");
    cleaned = cleaned.replace(/class="editor-paragraph"/g, "");
    cleaned = cleaned.replace(/<div>\s*<\/div>/g, "");
    cleaned = cleaned.replace(/<p>\s*<\/p>/g, "");
    cleaned = cleaned.replace(/<div>\s*<div>/g, "<div>");
    cleaned = cleaned.replace(/<\/div>\s*<\/div>/g, "</div>");
    return cleaned.trim();
}
function isValidHttpUrl(raw) {
    if (!raw)
        return false;
    try {
        var u = new URL(raw.trim());
        return u.protocol === "http:" || u.protocol === "https:";
    }
    catch (_a) {
        return false;
    }
}
export function savePmoItem(sp, pmoListId, id, draft) {
    return __awaiter(this, void 0, void 0, function () {
        function toMultiChoiceValue(v) {
            var arr = [];
            if (v == null) {
                arr = [];
            }
            else if (Array.isArray(v)) {
                arr = v;
            }
            else if (typeof v === "object" && Array.isArray(v.results)) {
                arr = v.results;
            }
            else if (typeof v === "string") {
                var s = v.trim();
                if (!s)
                    arr = [];
                else {
                    var parts = s.includes(";#") ? s.split(";#") : s.split(/[;,|]/);
                    arr = parts.map(function (x) { return x.trim(); }).filter(Boolean);
                }
            }
            else {
                arr = [String(v).trim()].filter(Boolean);
            }
            // ✅ MultiChoice לא מקבל null — רק results (גם אם ריק)
            return {
                "__metadata": { "type": "Collection(Edm.String)" },
                "results": arr
            };
        }
        var list, payload, tenderPhaseStr, isPhase23, MULTI_CHOICE_FIELDS, key, val, base, plus14, url, desc, e_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    console.log("savePmoItem 🌭");
                    list = sp.web.lists.getById(pmoListId);
                    payload = __assign({}, (draft || {}));
                    tenderPhaseStr = String(payload.TenderPhase || '').toLowerCase();
                    isPhase23 = tenderPhaseStr.indexOf('phase 2') > -1 ||
                        tenderPhaseStr.indexOf('phase 3') > -1;
                    if (isPhase23) {
                        console.log("🥔isPhase23 ", isPhase23);
                        delete payload.Assignedto;
                        delete payload.AssignedtoId; // אם זה People/Lookup – לפעמים נשמר כך
                        delete payload['Sub_x002d_Category'];
                        delete payload['Sub_x002d_CategoryId']; // אם זה Lookup
                        delete payload['SubCategory'];
                        delete payload['SubCategoryId'];
                    }
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
                    //delete payload.RFCorTcRFCasPublishedByNTaToBeFi;
                    delete payload.RFCorTcRFCasPublishedByNTaNew;
                    //payload.RevisedWordingFinalForPublicatio = cleanRichText(payload.RevisedWordingFinalForPublicatio);
                    console.log("payload.RevisedWordingFinalForPublicatio ", payload.RevisedWordingFinalForPublicatio, "\n____________________\n", cleanRichText(payload.RevisedWordingFinalForPublicatio));
                    console.log('🧾 payload sent to SharePoint:', payload);
                    console.log('🐶 dog in payload:', payload.dog);
                    MULTI_CHOICE_FIELDS = new Set([
                        "SubCategory",
                        "Sub_x002d_Category", // אם זה השם הפנימי אצלך
                    ]);
                    // GET ה־etag כדי SharePoint יאפשר עדכון
                    for (key in payload) {
                        val = payload[key];
                        // ✅ תיקון ספציפי ל-MultiChoice של SubCategory
                        if ((key === "SubCategory" || key === "Sub_x002d_Category") && (val == null || (Array.isArray(val) && val.length === 0) || (typeof val === "string" && !val.trim()))) {
                            delete payload[key]; // ✅ לא שולחים בכלל, כדי שלא ייכשל
                            continue;
                        }
                        if (MULTI_CHOICE_FIELDS.has(key)) {
                            payload[key] = toMultiChoiceValue(val);
                            continue;
                        }
                        if (!payload.RFCResponseLetterNo && payload.RFCResponseLetterNo !== 0) {
                            delete payload.RFCResponseLetterNo;
                        }
                        if (payload.DecisionDate) {
                            base = new Date(payload.DecisionDate);
                            plus14 = new Date(base.getTime() + 14 * 24 * 60 * 60 * 1000);
                            payload.DueDateCalculated = plus14.toISOString();
                        }
                        // Hyperlink field?
                        if (val && typeof val === "object" && ("Url" in val || "url" in val)) {
                            url = (val.Url || val.url || "").trim();
                            desc = (val.Description || val.description || url).trim();
                            if (!url) {
                                // אין URL → מאפסים
                                payload[key] = null;
                            }
                            else {
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
                    console.log("savePmoItem 🌭🌭");
                    console.log("SAVE PAYLOAD", JSON.stringify(payload, null, 2));
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, , 4]);
                    payload.__metadata = {
                        type: "SP.Data.PMO_x0020_decisionsListItem"
                    };
                    console.log("🦘 1");
                    // ⭐ UPDATE ללא metadata — רק REST
                    return [4 /*yield*/, spPost(list.items.getById(id), {
                            headers: {
                                "Accept": "application/json;odata=nometadata",
                                "Content-Type": "application/json;odata=verbose",
                                "IF-MATCH": "*",
                                "X-HTTP-Method": "MERGE"
                            },
                            body: JSON.stringify(payload)
                        })];
                case 2:
                    // ⭐ UPDATE ללא metadata — רק REST
                    _a.sent();
                    console.log("🦘 2");
                    return [3 /*break*/, 4];
                case 3:
                    e_1 = _a.sent();
                    console.log("savePMOITEM fall in the end with ", e_1);
                    return [3 /*break*/, 4];
                case 4:
                    console.log("savePmoItem 🌭🌭🌭");
                    return [4 /*yield*/, list.items.getById(id).select('*')()];
                case 5: 
                // 3. עדכון בפועל
                //await list.items.getById(id).update(payload);
                // 4. מחזירים את הפריט המעודכן
                return [2 /*return*/, _a.sent()];
            }
        });
    });
}
// =========================================================================
//   ==  הרשאות שדה ברמת PMO decisions לפי fieldPermission + קבוצות ==
//   ========================================================================= 
// 🔹 עמודות חדשות לשדות לפי M1/M2/M3 ברשימת fieldPermission
export var FP_COL_M1 = 'OData__x004d_1';
export var FP_COL_M2 = 'OData__x004d_2';
export var FP_COL_M3 = 'OData__x004d_3';
// טוען מפה: internalFieldName -> מערך תפקידי הרשאה (בחירות מ- WhoCanEdite) 
// טוען מפה: internalFieldName -> אובייקט הרשאות (ברירת מחדל + M1/M2/M3)
export function loadFieldPermissionMap(sp, fieldPermissionListTitle) {
    if (fieldPermissionListTitle === void 0) { fieldPermissionListTitle = FIELD_PERMISSION_LIST_ID; }
    return __awaiter(this, void 0, void 0, function () {
        var list, items, map, toChoices, i, it, internal, base, m1, m2, m3, entry, pushDistinctAll;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    console.log("😋1_ fieldPermissionListTitle ", fieldPermissionListTitle);
                    list = sp.web.lists.getById(fieldPermissionListTitle);
                    console.log("😋2_ FP_COL_INTERNAL ", FP_COL_INTERNAL, "FP_COL_WHO_CAN_EDIT ", FP_COL_WHO_CAN_EDIT);
                    return [4 /*yield*/, list.items.select('*')()];
                case 1:
                    items = _a.sent();
                    console.log("😋3_ ✏️ fieldPermission items: ", items);
                    map = {};
                    toChoices = function (raw) {
                        console.log("😋 4_");
                        if (!raw)
                            return [];
                        if (Array.isArray(raw))
                            return raw.map(function (x) { return String(x); });
                        if (raw && Array.isArray(raw.results))
                            return raw.results.map(function (x) { return String(x); });
                        if (typeof raw === 'string' && raw.trim())
                            return [raw.trim()];
                        return [];
                    };
                    for (i = 0; i < items.length; i++) {
                        it = items[i] || {};
                        internal = it[FP_COL_INTERNAL];
                        if (!internal)
                            continue;
                        base = toChoices(it[FP_COL_WHO_CAN_EDIT]);
                        m1 = toChoices(it[FP_COL_M1]);
                        m2 = toChoices(it[FP_COL_M2]);
                        m3 = toChoices(it[FP_COL_M3]);
                        entry = map[internal] || { default: [] };
                        pushDistinctAll = function (target, src) {
                            for (var j = 0; j < src.length; j++) {
                                var v = String(src[j]);
                                if (target.indexOf(v) === -1)
                                    target.push(v);
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
                    return [2 /*return*/, map];
            }
        });
    });
}
// ✅ ייצוא בשם מהמקור
export function fetchOrCreatePmoByIntegration(sp, pmoListTitle, integrationItemId, integrationLookupDisplayOrInternal) {
    return __awaiter(this, void 0, void 0, function () {
        var list, fields, e_2, lookupInternal, _i, _a, f, wantLc, found, created, item;
        var _b;
        return __generator(this, function (_c) {
            switch (_c.label) {
                case 0:
                    console.log("🔮1a");
                    list = sp.web.lists.getById('e5e8eaea-16db-49d3-ad7c-62f5a2bdd97a');
                    console.log("🔮2b");
                    _c.label = 1;
                case 1:
                    _c.trys.push([1, 3, , 4]);
                    return [4 /*yield*/, list.fields.select('Title', 'InternalName')()];
                case 2:
                    fields = _c.sent(); ///////////////////////////////////////////////
                    return [3 /*break*/, 4];
                case 3:
                    e_2 = _c.sent();
                    console.warn("🔮", e_2);
                    return [3 /*break*/, 4];
                case 4:
                    console.log("🔮3d");
                    lookupInternal = integrationLookupDisplayOrInternal;
                    console.log("🔮4r");
                    for (_i = 0, _a = fields; _i < _a.length; _i++) {
                        f = _a[_i];
                        wantLc = integrationLookupDisplayOrInternal.toLowerCase();
                        if (String(f.Title || '').toLowerCase() === wantLc ||
                            String(f.InternalName || '').toLowerCase() === wantLc) {
                            lookupInternal = f.InternalName;
                            break;
                        }
                    }
                    console.log("🔮5e");
                    return [4 /*yield*/, list.items.filter("".concat(lookupInternal, "Id eq ").concat(integrationItemId)).top(1)()];
                case 5:
                    found = _c.sent();
                    console.log("🔮6r");
                    if (found && found.length)
                        return [2 /*return*/, { item: found[0], isNew: false }];
                    console.log("🔮7g");
                    return [4 /*yield*/, list.items.add((_b = {}, _b["".concat(lookupInternal, "Id")] = integrationItemId, _b))];
                case 6:
                    created = _c.sent();
                    console.log("🔮8g");
                    return [4 /*yield*/, list.items.getById(created.data.Id).select('*')()];
                case 7:
                    item = _c.sent();
                    console.log("🔮9j");
                    return [2 /*return*/, { item: item, isNew: true }];
            }
        });
    });
}
var COL_INTEGRATION_TEAM = 'IntegrationTeam';
var COL_INTEGRATION_TEAM_LAWYER = 'IntegrationTeamLawyer';
var COL_FINANCIAL_ADVISOR_INTEGRATION_TEAM = 'FinancialAdvisorIntegrationTeam';
var COL_M1_TENDER_TEAM = 'M1TenderTeam';
var COL_LAWYER_M1_TENDER_TEAM = 'LawyerM1TenderTeam';
var COL_M2_TENDER_TEAM = 'M2TenderTeam';
var COL_LAWYER_M2_TENDER_TEAM = 'LawyerM2TenderTeam';
var COL_M3_TENDER_TEAM = 'M3TenderTeam';
var COL_LAWYER_M3_TENDER_TEAM = 'LawyerM3TenderTeam';
// שליפת כל המשתמשים ברשימת GeneralRoleDefinition לשלוש קבוצות + 9 עמודות טקסט
export function loadGeneralRoleUsers(sp) {
    return __awaiter(this, void 0, void 0, function () {
        var list, rows, fin, law, pmo, integrationTeam, integrationTeamLawyer, financialAdvisorIntegrationTeam, m1TenderTeam, lawyerM1TenderTeam, m2TenderTeam, lawyerM2TenderTeam, m3TenderTeam, lawyerM3TenderTeam, i, r, addPeople, addTextEmails;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    list = sp.web.lists.getById('b321608b-6405-4ad0-8676-11c7350fa7a4');
                    return [4 /*yield*/, list.items
                            .select('Id', "".concat(COL_FINANCIAL_ADVISOR, "/EMail"), "".concat(COL_LAWYER, "/EMail"), "".concat(COL_PMO_INTEGRATION_TEAM, "/EMail"), COL_INTEGRATION_TEAM, COL_INTEGRATION_TEAM_LAWYER, COL_FINANCIAL_ADVISOR_INTEGRATION_TEAM, COL_M1_TENDER_TEAM, COL_LAWYER_M1_TENDER_TEAM, COL_M2_TENDER_TEAM, COL_LAWYER_M2_TENDER_TEAM, COL_M3_TENDER_TEAM, COL_LAWYER_M3_TENDER_TEAM)
                            .expand(COL_FINANCIAL_ADVISOR, COL_LAWYER, COL_PMO_INTEGRATION_TEAM)()];
                case 1:
                    rows = _a.sent();
                    fin = [];
                    law = [];
                    pmo = [];
                    integrationTeam = [];
                    integrationTeamLawyer = [];
                    financialAdvisorIntegrationTeam = [];
                    m1TenderTeam = [];
                    lawyerM1TenderTeam = [];
                    m2TenderTeam = [];
                    lawyerM2TenderTeam = [];
                    m3TenderTeam = [];
                    lawyerM3TenderTeam = [];
                    for (i = 0; i < rows.length; i++) {
                        r = rows[i] || {};
                        addPeople = function (arr, target) {
                            if (!arr)
                                return;
                            var src = Array.isArray(arr) ? arr : (arr.results || []);
                            for (var j = 0; j < src.length; j++) {
                                var p = src[j];
                                var email = lc(p && (p.EMail || p.Email || p.UserPrincipalName));
                                if (email && target.indexOf(email) === -1)
                                    target.push(email);
                            }
                        };
                        addTextEmails = function (val, target) {
                            if (!val)
                                return;
                            var parts = String(val).split(/[;,\s]+/);
                            for (var k = 0; k < parts.length; k++) {
                                var email = lc(parts[k]);
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
                    return [2 /*return*/, {
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
                        }];
            }
        });
    });
}
export function loadTenderTeamUsersFromIntegration(sp, integrationItemId) {
    return __awaiter(this, void 0, void 0, function () {
        var it, raw, parts, emails, i, e;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    // שליפה של השדה כטקסט – ללא expand
                    console.log("🫙🔪🎡🎊🎀🎁🎉✨🧨🎆🎈 in loadTenderTeamUsersFromIntegration");
                    return [4 /*yield*/, integrationList(sp).items
                            .getById(integrationItemId)
                            .select('Id', COL_PMO_TENDER_TEAM_FROM_CREATER)()];
                case 1:
                    it = _a.sent();
                    console.log("it ", it);
                    raw = String((it === null || it === void 0 ? void 0 : it[COL_PMO_TENDER_TEAM_FROM_CREATER]) || '');
                    console.log("🪂 COL_PMO_TENDER_TEAM_FROM_CREATER ", COL_PMO_TENDER_TEAM_FROM_CREATER, "raw ", raw);
                    parts = raw.split(/[,;]|\r?\n/);
                    console.log("🎈parts ", parts);
                    emails = [];
                    for (i = 0; i < parts.length; i++) {
                        e = parts[i].trim().toLowerCase();
                        if (!e)
                            continue;
                        // ולידציה בסיסית של אימייל + מניעת כפילויות
                        if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e) && emails.indexOf(e) === -1) {
                            emails.push(e);
                        }
                    }
                    console.log("📧 emails ", emails);
                    return [2 /*return*/, emails];
            }
        });
    });
}
// החלטת עריכה לשדה ספציפי עבור משתמש נתון (email) /
// החלטת עריכה לשדה ספציפי עבור משתמש נתון (email) /
// כעת: לוקחת בחשבון גם OriginatingLineManager (M1/M2/M3)
export function canUserEditField(userEmailLc, internalName, fieldPermMap, roleUsers, originatingLineManager // 🔹 חדש: 'M1' | 'M2' | 'M3'
) {
    var entry = fieldPermMap[internalName];
    if (!entry) {
        console.log("🧨🛼internalName ", internalName, " entry ", entry);
        return false;
    } // אין הרשאות מוגדרות לשדה → קריאה בלבד
    // בחירת סט ה-"בחירות" לפי OriginatingLineManager:
    var olm = String(originatingLineManager || '').trim().toUpperCase();
    var choices = entry.default || [];
    if (olm === 'M1' && entry.M1 && entry.M1.length) {
        choices = entry.M1;
    }
    else if (olm === 'M2' && entry.M2 && entry.M2.length) {
        choices = entry.M2;
    }
    else if (olm === 'M3' && entry.M3 && entry.M3.length) {
        choices = entry.M3;
    }
    console.log("🧨 choices ", choices);
    if (!choices || !choices.length) {
        console.log("🧨🧨🧨🧨🧨🧨 !choices || !choices.length ");
        // אם גם ב-default וגם ב-M1/M2/M3 אין כלום → אין הרשאה
        return false;
    }
    // בניית איחוד משתמשים לפי הבחירות
    var allowEmails = [];
    var pushDistinct = function (src) {
        for (var i = 0; i < src.length; i++) {
            var e = src[i];
            if (allowEmails.indexOf(e) === -1)
                allowEmails.push(e);
        }
    };
    for (var i = 0; i < choices.length; i++) {
        var choice = choices[i];
        var cLc = lc(choice);
        if (cLc.indexOf(lc('צוות אינטגרציה')) > -1) {
            pushDistinct(roleUsers.IntegrationTeam);
            console.log("🧨🧨צוות אינטגרציה'");
        }
        else if (cLc.indexOf(lc('עו"ד צוות אינטגרציה')) > -1) {
            pushDistinct(roleUsers.IntegrationTeamLawyer);
            console.log("🧨🧨עוד צוות אינטגרציה");
        }
        else if (cLc.indexOf(lc('יועץ פיננסי צוות אינטגרציה')) > -1) {
            pushDistinct(roleUsers.FinancialAdvisorIntegrationTeam);
            console.log("🧨🧨יועץ פיננסי צוות אינטגרציה'");
        }
        else if (cLc.indexOf(lc('צוות מכרז M1')) > -1) {
            pushDistinct(roleUsers.M1TenderTeam);
            console.log("🧨🧨'צוות מכרז M1'");
        }
        else if (cLc.indexOf(lc('עו"ד צוות מכרז M1')) > -1) {
            pushDistinct(roleUsers.LawyerM1TenderTeam);
            console.log("🧨🧨עוד צוות מכרז M1");
        }
        else if (cLc.indexOf(lc('צוות מכרז M2')) > -1) {
            pushDistinct(roleUsers.M2TenderTeam);
            console.log("🧨🧨צוות מכרז M2 roleUsers.M2TenderTeam ", roleUsers.M2TenderTeam);
        }
        else if (cLc.indexOf(lc('עו"ד צוות מכרז M2')) > -1) {
            pushDistinct(roleUsers.LawyerM2TenderTeam);
            console.log("🧨🧨צוות מכרז עוד M2");
        }
        else if (cLc.indexOf(lc('צוות מכרז M3')) > -1) {
            pushDistinct(roleUsers.M3TenderTeam);
            console.log("🧨🧨צוות מכרז M3");
        }
        else if (cLc.indexOf(lc('עו"ד צוות מכרז M3')) > -1) {
            pushDistinct(roleUsers.LawyerM3TenderTeam);
            console.log("🧨🧨צוות מכרז עוד M3");
        }
        // כאן אפשר להוסיף עוד מיפויים אם יהיו בחירות נוספות בטבלה
    }
    console.log("🧨😶‍🌫️allowEmails  ", allowEmails);
    console.log(" return allowEmails.indexOf(userEmailLc) > -1 ", allowEmails.indexOf(userEmailLc) > -1);
    return allowEmails.indexOf(userEmailLc) > -1;
}
//# sourceMappingURL=data.js.map