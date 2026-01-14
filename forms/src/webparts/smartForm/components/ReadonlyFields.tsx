import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { DetailsList, IColumn } from '@fluentui/react/lib/DetailsList';
import { isSystemField } from '../shared/constants';

interface Props {
  item: any;
  /** מיפוי internal → display */
  internalToTitle?: Record<string, string>;
}

/** מציג את שדות ה-Integration כקריאה בלבד, עם שמות תצוגה וללא שדות מערכת */
const ReadonlyFields: React.FC<Props> = ({ item, internalToTitle = {} }) => {
  const rows = React.useMemo(() => {
    const pairs: { Field: string; Value: any }[] = [];
    Object.keys(item || {}).forEach((internal) => {
      if (isSystemField(internal)) return;

      const v = item[internal];
      if (v === null || typeof v === 'object') {
        // אפשר לפתוח רינדור מתקדם ל-Lookup/People/Choice במידת הצורך
        return;
      }
      const display = internalToTitle[internal] || internal; // תווית תצוגה
      pairs.push({ Field: display, Value: v });
    });
    return pairs;
  }, [item, internalToTitle]);

  const columns: IColumn[] = [
    { key: 'c1', name: 'שדה', fieldName: 'Field', minWidth: 180, isResizable: true },
    { key: 'c2', name: 'ערך', fieldName: 'Value', minWidth: 240, isResizable: true }
  ];

  return <Stack tokens={{ childrenGap: 6 }}><DetailsList items={rows} columns={columns} compact /></Stack>;
};

export default ReadonlyFields;
