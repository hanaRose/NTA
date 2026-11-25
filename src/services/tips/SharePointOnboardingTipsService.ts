import { SPHttpClient } from '@microsoft/sp-http';
import { IOnboardingTip } from './IOnboardingTip';
import { IOnboardingTipsService } from './IOnboardingTipsService';

interface IListField {
  Title: string;
  EntityPropertyName?: string;
}

interface IListItem {
  Id: unknown;
  Created: unknown;
  Title: unknown;
  [fieldName: string]: unknown;
}

interface IResolvedTip {
  tip: IOnboardingTip;
  createdAt: Date;
}

interface IListCollection<T> {
  value: T[];
  '@odata.nextLink'?: string;
  'odata.nextLink'?: string;
}

interface IHyperlinkValue {
  Url?: unknown;
}

const TIPS_LIST_TITLE: string = 'NTA_OB_Tips';
const CLUSTER_FIELD_TITLE: string = 'אשכול';
const DAY_FIELD_TITLE: string = 'יום במסע';
const STATUS_FIELD_TITLE: string = 'מצב פרסום';
const TEXT_FIELD_TITLE: string = 'תוכן הטיפ';
const LINK_FIELD_TITLE: string = 'קישור';
const LINK_LABEL_FIELD_TITLE: string = 'מלל הכפתור';
const SHOW_QR_FIELD_TITLE: string = 'הצגת QR';
const QR_IMAGE_FIELD_TITLE: string = 'תמונה';
const PUBLISHED_STATUS: string = 'פורסם';
const ACCEPT_HEADER: string = 'application/json;odata=nometadata';

function readRequiredFieldName(fields: Map<string, string>, title: string): string {
  const fieldName: string | undefined = fields.get(title);
  if (!fieldName) {
    throw new Error(`The SharePoint field "${title}" was not resolved.`);
  }
  return fieldName;
}

function readRequiredText(value: unknown, title: string): string {
  if (typeof value !== 'string' || !value.trim()) {
    throw new Error(`The ${title} field is empty or invalid.`);
  }
  return value.trim();
}

function readTipId(value: unknown): number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < 1) {
    throw new Error('A tip item has an invalid SharePoint ID.');
  }
  return value;
}

function readTipText(value: unknown): string {
  if (typeof value !== 'string') {
    throw new Error(`The ${TEXT_FIELD_TITLE} field contains an invalid value.`);
  }
  return value;
}

function readDay(value: unknown): number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < 0) {
    throw new Error(`The ${DAY_FIELD_TITLE} field must contain a non-negative whole number.`);
  }
  return value;
}

function readCreatedDate(value: unknown): Date {
  if (typeof value !== 'string') {
    throw new Error('A tip item has no valid SharePoint Created date.');
  }
  const createdAt: Date = new Date(value);
  if (Number.isNaN(createdAt.getTime())) {
    throw new Error('A tip item has an invalid SharePoint Created date.');
  }
  return createdAt;
}

function readLookupTitle(value: unknown): string {
  if (
    typeof value !== 'object' ||
    value === null ||
    !('Title' in value) ||
    typeof value.Title !== 'string' ||
    !value.Title.trim()
  ) {
    throw new Error(`The ${CLUSTER_FIELD_TITLE} lookup does not contain a title.`);
  }
  return value.Title.trim();
}

function readOptionalText(value: unknown, fieldTitle: string): string | undefined {
  if (value === null || value === undefined || value === '') {
    return undefined;
  }
  if (typeof value !== 'string') {
    throw new Error(`The ${fieldTitle} field contains an invalid value.`);
  }
  return value.trim() || undefined;
}

function readShowQr(value: unknown): boolean {
  if (value === null || value === undefined) {
    return false;
  }
  if (typeof value !== 'boolean') {
    throw new Error(`The ${SHOW_QR_FIELD_TITLE} field must contain a Yes/No value.`);
  }
  return value;
}

function resolveHttpUrl(rawUrl: string, baseUrl: string, fieldTitle: string): string {
  let url: URL;
  try {
    url = new URL(rawUrl.trim(), baseUrl);
  } catch {
    throw new Error(`The ${fieldTitle} field contains an invalid URL.`);
  }
  if (url.protocol !== 'http:' && url.protocol !== 'https:') {
    throw new Error(`The ${fieldTitle} field URL must use HTTP or HTTPS.`);
  }
  return url.toString();
}

function readOptionalTipLink(value: unknown, baseUrl: string): string | undefined {
  if (value === null || value === undefined || value === '') {
    return undefined;
  }
  const rawUrl: unknown = typeof value === 'object' && value !== null && 'Url' in value
    ? (value as IHyperlinkValue).Url
    : value;
  if (typeof rawUrl !== 'string' || !rawUrl.trim()) {
    throw new Error(`The ${LINK_FIELD_TITLE} field contains an invalid hyperlink.`);
  }
  return resolveHttpUrl(rawUrl, baseUrl, LINK_FIELD_TITLE);
}

function readQrImageUrl(value: unknown, baseUrl: string): string | undefined {
  if (value === null || value === undefined || value === '') {
    return undefined;
  }

  let imageValue: unknown = value;
  if (typeof imageValue === 'string') {
    const rawValue: string = imageValue.trim();
    if (!rawValue) {
      return undefined;
    }
    if (rawValue.charAt(0) === '{') {
      try {
        imageValue = JSON.parse(rawValue) as unknown;
      } catch {
        throw new Error(`The ${QR_IMAGE_FIELD_TITLE} field contains invalid image data.`);
      }
    } else {
      return resolveHttpUrl(rawValue, baseUrl, QR_IMAGE_FIELD_TITLE);
    }
  }

  if (typeof imageValue !== 'object' || imageValue === null || Array.isArray(imageValue)) {
    throw new Error(`The ${QR_IMAGE_FIELD_TITLE} field contains an unsupported image value.`);
  }

  const image = imageValue as Record<string, unknown>;
  const serverRelativeUrl: unknown = image.serverRelativeUrl;
  if (typeof serverRelativeUrl === 'string' && serverRelativeUrl.trim()) {
    const serverUrl: unknown = image.serverUrl;
    const imageBaseUrl: string = typeof serverUrl === 'string' ? serverUrl : baseUrl;
    return resolveHttpUrl(serverRelativeUrl, imageBaseUrl, QR_IMAGE_FIELD_TITLE);
  }

  const imageUrl: unknown = image.Url !== undefined ? image.Url : image.url;
  if (typeof imageUrl === 'string' && imageUrl.trim()) {
    return resolveHttpUrl(imageUrl, baseUrl, QR_IMAGE_FIELD_TITLE);
  }

  if (image.nativeFile !== undefined) {
    return readQrImageUrl(image.nativeFile, baseUrl);
  }

  throw new Error(`The ${QR_IMAGE_FIELD_TITLE} field does not contain an image URL.`);
}

export default class SharePointOnboardingTipsService implements IOnboardingTipsService {
  private readonly _webUrl: string;

  public constructor(
    private readonly _spHttpClient: SPHttpClient,
    webUrl: string
  ) {
    this._webUrl = webUrl.replace(/\/$/, '');
  }

  public async getPublishedTips(): Promise<IOnboardingTip[]> {
    const fields: Map<string, string> = await this._getFieldNames();
    const clusterField: string = readRequiredFieldName(fields, CLUSTER_FIELD_TITLE);
    const dayField: string = readRequiredFieldName(fields, DAY_FIELD_TITLE);
    const statusField: string = readRequiredFieldName(fields, STATUS_FIELD_TITLE);
    const textField: string = readRequiredFieldName(fields, TEXT_FIELD_TITLE);
    const linkField: string = readRequiredFieldName(fields, LINK_FIELD_TITLE);
    const linkLabelField: string = readRequiredFieldName(fields, LINK_LABEL_FIELD_TITLE);
    const showQrField: string = readRequiredFieldName(fields, SHOW_QR_FIELD_TITLE);
    const qrImageField: string = readRequiredFieldName(fields, QR_IMAGE_FIELD_TITLE);
    const selectedFields: string[] = [
      'Id',
      'Title',
      'Created',
      `${clusterField}/Title`,
      dayField,
      statusField,
      textField,
      linkField,
      linkLabelField,
      showQrField,
      qrImageField
    ];
    const filter: string = encodeURIComponent(`${statusField} eq '${PUBLISHED_STATUS}'`);
    const query: string =
      `?$select=${encodeURIComponent(selectedFields.join(','))}` +
      `&$expand=${encodeURIComponent(clusterField)}` +
      `&$filter=${filter}&$orderby=Created desc&$top=5000`;
    const items: IListItem[] = await this._getAllItems<IListItem>(
      `${this._listItemsUrl()}${query}`,
      TIPS_LIST_TITLE
    );

    const resolvedTips: IResolvedTip[] = items
      .filter((item: IListItem) => item[statusField] === PUBLISHED_STATUS)
      .map((item: IListItem) => {
        const showQr: boolean = readShowQr(item[showQrField]);
        const qrImageUrl: string | undefined = showQr
          ? readQrImageUrl(item[qrImageField], this._webUrl)
          : undefined;
        const linkLabel: string | undefined = readOptionalText(
          item[linkLabelField],
          LINK_LABEL_FIELD_TITLE
        );
        const link: string | undefined = readOptionalTipLink(item[linkField], this._webUrl);

        return {
          tip: {
            id: readTipId(item.Id),
            title: readRequiredText(item.Title, 'Title'),
            text: readTipText(item[textField]),
            day: readDay(item[dayField]),
            category: readLookupTitle(item[clusterField]),
            link,
            linkLabel,
            showQr,
            qrImageUrl
          },
          createdAt: readCreatedDate(item.Created)
        };
      })
      .sort((first: IResolvedTip, second: IResolvedTip) =>
        second.createdAt.getTime() - first.createdAt.getTime() ||
        second.tip.id - first.tip.id
      );
    return resolvedTips.map((resolvedTip: IResolvedTip) => resolvedTip.tip);
  }

  private async _getFieldNames(): Promise<Map<string, string>> {
    const url: string = `${this._listUrl()}/fields?$select=Title,EntityPropertyName&$top=5000`;
    const fields: IListField[] = await this._getAllItems<IListField>(
      url,
      `${TIPS_LIST_TITLE} fields`
    );
    const titles: string[] = [
      CLUSTER_FIELD_TITLE,
      DAY_FIELD_TITLE,
      STATUS_FIELD_TITLE,
      TEXT_FIELD_TITLE,
      LINK_FIELD_TITLE,
      LINK_LABEL_FIELD_TITLE,
      SHOW_QR_FIELD_TITLE,
      QR_IMAGE_FIELD_TITLE
    ];
    const fieldNames: Map<string, string> = new Map<string, string>();

    titles.forEach((title: string) => {
      const matches: IListField[] = fields.filter((field: IListField) => field.Title === title);
      if (matches.length !== 1 || !matches[0].EntityPropertyName) {
        throw new Error(`Could not uniquely resolve the "${title}" field in ${TIPS_LIST_TITLE}.`);
      }
      fieldNames.set(title, matches[0].EntityPropertyName);
    });

    return fieldNames;
  }

  private _listUrl(): string {
    const escapedTitle: string = TIPS_LIST_TITLE.replace(/'/g, "''");
    return `${this._webUrl}/_api/web/lists/GetByTitle('${escapedTitle}')`;
  }

  private _listItemsUrl(): string {
    return `${this._listUrl()}/items`;
  }

  private async _getAllItems<T>(url: string, sourceName: string): Promise<T[]> {
    const items: T[] = [];
    let nextUrl: string | undefined = url;

    while (nextUrl) {
      const response = await this._spHttpClient.get(
        nextUrl,
        SPHttpClient.configurations.v1,
        { headers: { Accept: ACCEPT_HEADER } }
      );
      if (!response.ok) {
        throw new Error(`Could not read ${sourceName} (${response.status} ${response.statusText}).`);
      }

      const result = await response.json() as IListCollection<T>;
      if (!Array.isArray(result.value)) {
        throw new Error(`The ${sourceName} response did not contain a valid item collection.`);
      }
      items.push(...result.value);
      nextUrl = result['@odata.nextLink'] || result['odata.nextLink'];
    }

    return items;
  }
}
