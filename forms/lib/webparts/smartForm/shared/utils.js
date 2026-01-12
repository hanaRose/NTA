/** ניסיון לזהות את ה-ID של הפריט מתוך ה-Query String או הקונטקסט של הדף */
export function getQueryId(context) {
    var _a, _b;
    var qs = new URLSearchParams(window.location.search);
    var str = qs.get('iid');
    if (str && /^\d+$/.test(str))
        return Number(str);
    var pageItemId = (_b = (_a = context === null || context === void 0 ? void 0 : context.pageContext) === null || _a === void 0 ? void 0 : _a.listItem) === null || _b === void 0 ? void 0 : _b.id;
    if (pageItemId)
        return Number(pageItemId);
    return null;
}
//# sourceMappingURL=utils.js.map