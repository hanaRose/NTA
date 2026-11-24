import { SPHttpClient } from '@microsoft/sp-http';
import { IOnboardingTaskMetrics } from './IOnboardingTaskMetrics';
import { IOnboardingTaskMetricsService } from './IOnboardingTaskMetricsService';
import {
  calculateOnboardingTaskMetrics,
  ICompletedTaskReference,
  ITaskMetricCandidate
} from './calculateOnboardingTaskMetrics';

interface IListField {
  Title: string;
  EntityPropertyName?: string;
}

interface ISharePointCollection<T> {
  value: T[];
  '@odata.nextLink'?: string;
  'odata.nextLink'?: string;
}

interface ITaskListItem {
  Id: number;
  [fieldName: string]: unknown;
}

interface ITaskCompletionListItem {
  Id: number;
  [fieldName: string]: unknown;
}

const TASKS_LIST_TITLE: string = 'NTA_OB_Tasks';
const COMPLETIONS_LIST_TITLE: string = 'NTA_OB_TaskCompletions';
const PUBLISHED_STATUS: string = 'פורסם';
const TASK_DAY_FIELD_TITLE: string = 'יום במסע';
const TASK_STATUS_FIELD_TITLE: string = 'מצב פרסום';
const TASK_DURATION_FIELD_TITLE: string = 'משך';
const TASK_CLUSTER_FIELD_TITLE: string = 'אשכול';
const COMPLETION_JOURNEY_FIELD_TITLE: string = 'Journey';
const COMPLETION_TASK_FIELD_TITLE: string = 'Task';
const SHAREPOINT_ACCEPT_HEADER: string = 'application/json;odata=nometadata';

function readOptionalNumber(value: unknown, fieldTitle: string): number | undefined {
  if (value === null || value === undefined) {
    return undefined;
  }
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw new Error(`The ${fieldTitle} field contains a non-numeric value.`);
  }
  return value;
}

function readLookupTitle(value: unknown, fieldTitle: string): string {
  if (
    typeof value !== 'object' ||
    value === null ||
    !('Title' in value) ||
    typeof value.Title !== 'string' ||
    !value.Title.trim()
  ) {
    throw new Error(`The ${fieldTitle} lookup does not contain a title.`);
  }
  return value.Title;
}

export default class SharePointOnboardingTaskMetricsService implements IOnboardingTaskMetricsService {
  private readonly _webUrl: string;

  public constructor(
    private readonly _spHttpClient: SPHttpClient,
    webUrl: string
  ) {
    this._webUrl = webUrl.replace(/\/$/, '');
  }

  public async getForJourney(
    journeyId: number,
    currentDay: number
  ): Promise<IOnboardingTaskMetrics> {
    if (!Number.isInteger(journeyId) || journeyId < 1) {
      throw new Error('The current SharePoint journey ID is invalid.');
    }
    if (!Number.isFinite(currentDay) || currentDay < 1) {
      throw new Error('The current onboarding journey day is invalid.');
    }

    const [taskFields, completionFields] = await Promise.all([
      this._getFieldNames(TASKS_LIST_TITLE, [
        TASK_DAY_FIELD_TITLE,
        TASK_STATUS_FIELD_TITLE,
        TASK_DURATION_FIELD_TITLE,
        TASK_CLUSTER_FIELD_TITLE
      ]),
      this._getFieldNames(COMPLETIONS_LIST_TITLE, [
        COMPLETION_JOURNEY_FIELD_TITLE,
        COMPLETION_TASK_FIELD_TITLE
      ])
    ]);

    const taskDayField: string = this._getRequiredFieldName(taskFields, TASK_DAY_FIELD_TITLE);
    const taskStatusField: string = this._getRequiredFieldName(taskFields, TASK_STATUS_FIELD_TITLE);
    const taskDurationField: string = this._getRequiredFieldName(taskFields, TASK_DURATION_FIELD_TITLE);
    const taskClusterField: string = this._getRequiredFieldName(taskFields, TASK_CLUSTER_FIELD_TITLE);
    const journeyLookupIdField: string =
      `${this._getRequiredFieldName(completionFields, COMPLETION_JOURNEY_FIELD_TITLE)}Id`;
    const taskLookupIdField: string =
      `${this._getRequiredFieldName(completionFields, COMPLETION_TASK_FIELD_TITLE)}Id`;

    const taskSelect: string = encodeURIComponent(
      ['Id', taskDayField, taskStatusField, taskDurationField, `${taskClusterField}/Title`].join(',')
    );
    const tasksUrl: string = `${this._listItemsUrl(TASKS_LIST_TITLE)}?` +
      `$select=${taskSelect}&$expand=${encodeURIComponent(taskClusterField)}&$top=5000`;
    const completionSelect: string = encodeURIComponent(
      ['Id', journeyLookupIdField, taskLookupIdField].join(',')
    );
    const completionFilter: string = encodeURIComponent(`${journeyLookupIdField} eq ${journeyId}`);
    const completionsUrl: string = `${this._listItemsUrl(COMPLETIONS_LIST_TITLE)}?` +
      `$select=${completionSelect}&$filter=${completionFilter}&$top=5000`;

    const [allTasks, journeyCompletions] = await Promise.all([
      this._getAllItems<ITaskListItem>(tasksUrl, TASKS_LIST_TITLE),
      this._getAllItems<ITaskCompletionListItem>(completionsUrl, COMPLETIONS_LIST_TITLE)
    ]);

    const tasks: ITaskMetricCandidate[] = allTasks.map((task: ITaskListItem) => ({
      id: task.Id,
      isPublished: task[taskStatusField] === PUBLISHED_STATUS,
      clusterName: readLookupTitle(task[taskClusterField], TASK_CLUSTER_FIELD_TITLE),
      journeyDay: readOptionalNumber(task[taskDayField], TASK_DAY_FIELD_TITLE),
      duration: readOptionalNumber(task[taskDurationField], TASK_DURATION_FIELD_TITLE)
    }));
    const completions: ICompletedTaskReference[] = journeyCompletions.map(
      (completion: ITaskCompletionListItem) => {
        const taskId: unknown = completion[taskLookupIdField];
        if (typeof taskId !== 'number' || !Number.isInteger(taskId) || taskId < 1) {
          throw new Error(
            `A ${COMPLETIONS_LIST_TITLE} item for this journey has an invalid Task lookup.`
          );
        }
        return { taskId };
      }
    );

    return calculateOnboardingTaskMetrics(tasks, completions, currentDay);
  }

  private async _getFieldNames(listTitle: string, fieldTitles: string[]): Promise<Map<string, string>> {
    const url: string = `${this._listUrl(listTitle)}/fields?$select=Title,EntityPropertyName&$top=5000`;
    const fields: IListField[] = await this._getAllItems<IListField>(url, `${listTitle} fields`);
    const fieldNames: Map<string, string> = new Map<string, string>();

    fieldTitles.forEach((fieldTitle: string) => {
      const matches: IListField[] = fields.filter((field: IListField) => field.Title === fieldTitle);
      if (matches.length !== 1 || !matches[0].EntityPropertyName) {
        throw new Error(
          `Could not uniquely resolve the "${fieldTitle}" field in the ${listTitle} list.`
        );
      }
      fieldNames.set(fieldTitle, matches[0].EntityPropertyName);
    });

    return fieldNames;
  }

  private _getRequiredFieldName(fields: Map<string, string>, title: string): string {
    const fieldName: string | undefined = fields.get(title);
    if (!fieldName) {
      throw new Error(`The SharePoint field "${title}" was not resolved.`);
    }
    return fieldName;
  }

  private _listUrl(listTitle: string): string {
    const escapedTitle: string = listTitle.replace(/'/g, "''");
    return `${this._webUrl}/_api/web/lists/GetByTitle('${escapedTitle}')`;
  }

  private _listItemsUrl(listTitle: string): string {
    return `${this._listUrl(listTitle)}/items`;
  }

  private async _getAllItems<T>(url: string, sourceName: string): Promise<T[]> {
    const items: T[] = [];
    let nextUrl: string | undefined = url;

    while (nextUrl) {
      const response = await this._spHttpClient.get(
        nextUrl,
        SPHttpClient.configurations.v1,
        { headers: { Accept: SHAREPOINT_ACCEPT_HEADER } }
      );

      if (!response.ok) {
        throw new Error(
          `Could not read ${sourceName} (${response.status} ${response.statusText}).`
        );
      }

      const result = await response.json() as ISharePointCollection<T>;
      if (!Array.isArray(result.value)) {
        throw new Error(`The ${sourceName} response did not contain a valid item collection.`);
      }

      items.push(...result.value);
      nextUrl = result['@odata.nextLink'] || result['odata.nextLink'];
    }

    return items;
  }
}
