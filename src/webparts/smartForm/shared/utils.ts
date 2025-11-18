import type { WebPartContext } from '@microsoft/sp-webpart-base';

/** ניסיון לזהות את ה-ID של הפריט מתוך ה-Query String או הקונטקסט של הדף */
export function getQueryId(context: WebPartContext): number | null {
  const qs = new URLSearchParams(window.location.search);
  const str = qs.get('iid');
  if (str && /^\d+$/.test(str)) return Number(str);

  const pageItemId = (context as any)?.pageContext?.listItem?.id;
  if (pageItemId) return Number(pageItemId);

  return null;
}
