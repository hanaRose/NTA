import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { DetailsList } from '@fluentui/react/lib/DetailsList';
import { isSystemField } from '../shared/constants';
/** מציג את שדות ה-Integration כקריאה בלבד, עם שמות תצוגה וללא שדות מערכת */
var ReadonlyFields = function (_a) {
    var item = _a.item, _b = _a.internalToTitle, internalToTitle = _b === void 0 ? {} : _b;
    var rows = React.useMemo(function () {
        var pairs = [];
        Object.keys(item || {}).forEach(function (internal) {
            if (isSystemField(internal))
                return;
            var v = item[internal];
            if (v === null || typeof v === 'object') {
                // אפשר לפתוח רינדור מתקדם ל-Lookup/People/Choice במידת הצורך
                return;
            }
            var display = internalToTitle[internal] || internal; // תווית תצוגה
            pairs.push({ Field: display, Value: v });
        });
        return pairs;
    }, [item, internalToTitle]);
    var columns = [
        { key: 'c1', name: 'שדה', fieldName: 'Field', minWidth: 180, isResizable: true },
        { key: 'c2', name: 'ערך', fieldName: 'Value', minWidth: 240, isResizable: true }
    ];
    return React.createElement(Stack, { tokens: { childrenGap: 6 } },
        React.createElement(DetailsList, { items: rows, columns: columns, compact: true }));
};
export default ReadonlyFields;
//# sourceMappingURL=ReadonlyFields.js.map