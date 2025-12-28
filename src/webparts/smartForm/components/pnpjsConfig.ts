//src\webparts\smartForm\components\pnpjsConfig.ts
import { spfi, SPFI } from '@pnp/sp';
import { SPFx } from '@pnp/sp';
import type { WebPartContext } from '@microsoft/sp-webpart-base';

let _sp: SPFI | null = null;

export const getSP = (context: WebPartContext): SPFI => {
  // חובה להעביר context בקריאה *הראשונה*
  if (!_sp) {
    if (!context) {
      throw new Error('getSP: context is required on first call');
    }
    _sp = spfi().using(SPFx(context));
  }

  return _sp;
};