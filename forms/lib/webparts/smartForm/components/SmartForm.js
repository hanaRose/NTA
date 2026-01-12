//src\webparts\smartForm\components\SmartForm.tsx
import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { DetailsList } from '@fluentui/react/lib/DetailsList';
/** מציג את כל השדות כקריאה בלבד (מלבד אלה שכבר הוחרגו בשאיבה) */
var ReadonlyFields = function (_a) {
    var item = _a.item;
    var rows = React.useMemo(function () {
        var pairs = [];
        Object.keys(item || {}).forEach(function (k) {
            var v = item[k];
            if (v === null || typeof v === 'object')
                return; // skip מורכבים; או פתח לפי צורך
            pairs.push({ Field: k, Value: v });
        });
        return pairs;
    }, [item]);
    var columns = [
        { key: 'c1', name: 'שדה', fieldName: 'Field', minWidth: 120, isResizable: true },
        { key: 'c2', name: 'ערך', fieldName: 'Value', minWidth: 200, isResizable: true }
    ];
    return (React.createElement(Stack, { tokens: { childrenGap: 6 } },
        React.createElement(DetailsList, { items: rows, columns: columns, compact: true }),
        React.createElement(Text, { variant: "small" }, "* Values \u200B\u200Bare extracted directly from Integration")));
};
export default ReadonlyFields;
//# sourceMappingURL=SmartForm.js.map