import * as React from 'react';
interface Props {
    item: any;
    onChange: (internalName: string, value: any) => void;
    fieldOrder?: string[];
    hideFields?: string[];
    internalToTitle?: Record<string, string>;
    fieldInfoMap?: Record<string, any>;
    canEdit?: (internalName: string) => boolean;
    placeholderMap?: Partial<Record<string, string>>;
    choiceOverrides?: Record<string, {
        key: string;
        text: string;
    }[]>;
    tenderPhase?: string;
    requiredMap?: Record<string, boolean>;
    errorMap?: Record<string, string>;
    labelOverrides?: Record<string, string>;
}
export declare const formatDateDDMMYYYY: (d?: Date | string | null) => string;
declare const EditableFields: React.FC<Props>;
export default EditableFields;
//# sourceMappingURL=EditableFields.d.ts.map