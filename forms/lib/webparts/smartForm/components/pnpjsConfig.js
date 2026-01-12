//src\webparts\smartForm\components\pnpjsConfig.ts
import { spfi } from '@pnp/sp';
import { SPFx } from '@pnp/sp';
var _sp = null;
export var getSP = function (context) {
    // חובה להעביר context בקריאה *הראשונה*
    if (!_sp) {
        if (!context) {
            throw new Error('getSP: context is required on first call');
        }
        _sp = spfi().using(SPFx(context));
    }
    return _sp;
};
//# sourceMappingURL=pnpjsConfig.js.map