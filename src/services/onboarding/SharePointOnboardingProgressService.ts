import { SPHttpClient } from '@microsoft/sp-http';
import { IOnboardingProgress } from './IOnboardingProgress';
import { IOnboardingProgressService } from './IOnboardingProgressService';

interface IProgressListItem {
  Id: number;
  StartDate?: string;
}

interface IProgressListResponse {
  value: IProgressListItem[];
}

const LIST_TITLE: string = 'NTA_OB_Journeys';
const DAY_IN_MILLISECONDS: number = 24 * 60 * 60 * 1000;

export default class SharePointOnboardingProgressService implements IOnboardingProgressService {
  private readonly _webUrl: string;

  public constructor(
    private readonly _spHttpClient: SPHttpClient,
    webUrl: string
  ) {
    this._webUrl = webUrl.replace(/\/$/, '');
  }

  public async getForUser(userEmail: string): Promise<IOnboardingProgress> {
    const normalizedEmail: string = userEmail.trim();
    if (!normalizedEmail) {
      throw new Error('The current SharePoint user email is unavailable.');
    }

    const escapedEmail: string = normalizedEmail.replace(/'/g, "''");
    const filter: string = encodeURIComponent(`SPUser/EMail eq '${escapedEmail}'`);
    const url: string = `${this._webUrl}/_api/web/lists/GetByTitle('${LIST_TITLE}')/items` +
      `?$select=Id,StartDate,SPUser/EMail&$expand=SPUser&$filter=${filter}&$top=2`;
    const response = await this._spHttpClient.get(
      url,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );

    if (!response.ok) {
      throw new Error(
        `Could not read the ${LIST_TITLE} list (${response.status} ${response.statusText}).`
      );
    }

    const result = await response.json() as IProgressListResponse;
    if (!result.value || result.value.length === 0) {
      throw new Error(`No ${LIST_TITLE} record was found for the current SharePoint user.`);
    }
    if (result.value.length > 1) {
      throw new Error(`More than one ${LIST_TITLE} record exists for the current SharePoint user.`);
    }

    const startDateValue: string | undefined = result.value[0].StartDate;
    if (!startDateValue) {
      throw new Error(`The ${LIST_TITLE} record has no StartDate value.`);
    }

    const startDate: Date = new Date(startDateValue);
    if (Number.isNaN(startDate.getTime())) {
      throw new Error(`The ${LIST_TITLE} StartDate value is invalid.`);
    }

    const today: Date = new Date();
    const todayUtc: number = Date.UTC(today.getFullYear(), today.getMonth(), today.getDate());
    const startDateUtc: number = Date.UTC(
      startDate.getUTCFullYear(),
      startDate.getUTCMonth(),
      startDate.getUTCDate()
    );
    const elapsedDays: number = Math.floor((todayUtc - startDateUtc) / DAY_IN_MILLISECONDS);
    const currentDay: number = Math.max(1, elapsedDays + 1);

    return { journeyId: result.value[0].Id, startDate, currentDay };
  }
}
