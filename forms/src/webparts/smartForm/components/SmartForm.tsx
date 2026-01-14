//src\webparts\smartForm\components\SmartForm.tsx
import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { DetailsList, IColumn } from '@fluentui/react/lib/DetailsList';

interface Props {
  item: any;
}

/** מציג את כל השדות כקריאה בלבד (מלבד אלה שכבר הוחרגו בשאיבה) */
const ReadonlyFields: React.FC<Props> = ({ item }) => {
  const rows = React.useMemo(() => {
    const pairs: { Field: string; Value: any }[] = [];
    Object.keys(item || {}).forEach(k => {
      const v = item[k];
      if (v === null || typeof v === 'object') return; // skip מורכבים; או פתח לפי צורך
      pairs.push({ Field: k, Value: v });
    });
    return pairs;
  }, [item]);

  const columns: IColumn[] = [
    { key: 'c1', name: 'שדה', fieldName: 'Field', minWidth: 120, isResizable: true },
    { key: 'c2', name: 'ערך', fieldName: 'Value', minWidth: 200, isResizable: true }
  ];

  return (
    <Stack tokens={{ childrenGap: 6 }}>
      <DetailsList items={rows} columns={columns} compact />
      <Text variant="small">* Values ​​are extracted directly from Integration</Text>
    </Stack>
  );
};

export default ReadonlyFields;
