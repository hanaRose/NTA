import { SPHttpClient } from '@microsoft/sp-http';
import { IOnboardingTask } from './IOnboardingTask';
import { IOnboardingTaskMetrics } from './IOnboardingTaskMetrics';
import { IOnboardingTaskMetricsService } from './IOnboardingTaskMetricsService';
import { IOnboardingTasksService } from './IOnboardingTasksService';
import { readOnboardingTaskContentUrl } from './readOnboardingTaskContentUrl';
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
  Title?: unknown;
  Created?: unknown;
  [fieldName: string]: unknown;
}

interface ITaskCompletionListItem {
  Id: number;
  [fieldName: string]: unknown;
}

interface ICreatedTaskCompletion {
  Id?: unknown;
  ID?: unknown;
}

interface IOnboardingJourneyData {
  taskDayField: string;
  taskStatusField: string;
  taskDurationField: string;
  taskClusterField: string;
  taskContentTypeField?: string;
  taskLookupIdField: string;
  tasks: ITaskListItem[];
  completions: ITaskCompletionListItem[];
}

const TASKS_LIST_TITLE: string = 'NTA_OB_Tasks';
const COMPLETIONS_LIST_TITLE: string = 'NTA_OB_TaskCompletions';
const PUBLISHED_STATUS: string = 'פורסם';
const TASK_DAY_FIELD_TITLE: string = 'יום במסע';
const TASK_STATUS_FIELD_TITLE: string = 'מצב פרסום';
const TASK_DURATION_FIELD_TITLE: string = 'משך';
const TASK_CLUSTER_FIELD_TITLE: string = 'אשכול';
const TASK_CONTENT_TYPE_FIELD_TITLE: string = 'סיווג התוכן';
const TASK_CONTENT_URL_FIELD_NAME: string = 'ContentUrl';
const COMPLETION_JOURNEY_FIELD_TITLE: string = 'Journey';
const COMPLETION_TASK_FIELD_TITLE: string = 'Task';
const COMPLETION_JOURNEY_DAY_FIELD_TITLE: string = 'JourneyDayAtComplete';
const COMPLETION_DATE_FIELD_TITLE: string = 'CompletionDate';
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

function readRequiredNumber(value: unknown, fieldTitle: string): number {
  const parsedValue: number | undefined = readOptionalNumber(value, fieldTitle);
  if (parsedValue === undefined || !Number.isInteger(parsedValue) || parsedValue < 0) {
    throw new Error(`The ${fieldTitle} field must contain a non-negative whole number.`);
  }
  return parsedValue;
}

function readRequiredText(value: unknown, fieldTitle: string): string {
  if (typeof value !== 'string' || !value.trim()) {
    throw new Error(`The ${fieldTitle} field is empty or invalid.`);
  }
  return value.trim();
}

function readCreatedDate(value: unknown): Date {
  if (typeof value !== 'string') {
    throw new Error('A task item has no valid SharePoint Created date.');
  }
  const createdAt: Date = new Date(value);
  if (Number.isNaN(createdAt.getTime())) {
    throw new Error('A task item has an invalid SharePoint Created date.');
  }
  return createdAt;
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

export default class SharePointOnboardingTaskMetricsService implements
  IOnboardingTaskMetricsService,
  IOnboardingTasksService {
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
    if (!Number.isFinite(currentDay) || currentDay < 1) {
      throw new Error('The current onboarding journey day is invalid.');
    }

    const journeyData: IOnboardingJourneyData = await this._getJourneyData(journeyId);
    const tasks: ITaskMetricCandidate[] = journeyData.tasks.map((task: ITaskListItem) => ({
      id: task.Id,
      isPublished: task[journeyData.taskStatusField] === PUBLISHED_STATUS,
      clusterName: readLookupTitle(task[journeyData.taskClusterField], TASK_CLUSTER_FIELD_TITLE),
      journeyDay: readOptionalNumber(task[journeyData.taskDayField], TASK_DAY_FIELD_TITLE),
      duration: readOptionalNumber(task[journeyData.taskDurationField], TASK_DURATION_FIELD_TITLE)
    }));
    const completedTaskIds: Set<number> = this._readCompletedTaskIds(
      journeyData.completions,
      journeyData.taskLookupIdField
    );
    const completions: ICompletedTaskReference[] = Array.from(completedTaskIds).map(
      (taskId: number) => ({ taskId })
    );

    return calculateOnboardingTaskMetrics(tasks, completions, currentDay);
  }

  public async getIncompleteForJourney(journeyId: number): Promise<IOnboardingTask[]> {
    const journeyData: IOnboardingJourneyData = await this._getJourneyData(journeyId, true);
    if (!journeyData.taskContentTypeField) {
      throw new Error(`The SharePoint field "${TASK_CONTENT_TYPE_FIELD_TITLE}" was not resolved.`);
    }
    const taskContentTypeField: string = journeyData.taskContentTypeField;
    const completedTaskIds: Set<number> = this._readCompletedTaskIds(
      journeyData.completions,
      journeyData.taskLookupIdField
    );

    return journeyData.tasks
      .filter((task: ITaskListItem) =>
        task[journeyData.taskStatusField] === PUBLISHED_STATUS &&
        !completedTaskIds.has(task.Id)
      )
      .map((task: ITaskListItem) => {
        const duration: number | undefined = readOptionalNumber(
          task[journeyData.taskDurationField],
          TASK_DURATION_FIELD_TITLE
        );
        if (duration !== undefined && duration < 0) {
          throw new Error('The משך field cannot contain a negative value.');
        }

        return {
          id: task.Id,
          title: readRequiredText(task.Title, 'Title'),
          taskContentType: readRequiredText(
            task[taskContentTypeField],
            TASK_CONTENT_TYPE_FIELD_TITLE
          ),
          url: readOnboardingTaskContentUrl(
            task[TASK_CONTENT_URL_FIELD_NAME],
            this._webUrl
          ),
          journeyDay: readRequiredNumber(task[journeyData.taskDayField], TASK_DAY_FIELD_TITLE),
          duration,
          clusterName: readLookupTitle(
            task[journeyData.taskClusterField],
            TASK_CLUSTER_FIELD_TITLE
          ),
          createdAt: readCreatedDate(task.Created)
        };
      });
  }

  public async createCompletion(
    journeyId: number,
    taskId: number,
    taskTitle: string,
    journeyDay: number
  ): Promise<number> {
    this._validateId(journeyId, 'journey');
    this._validateId(taskId, 'task');
    if (!Number.isInteger(journeyDay) || journeyDay < 0) {
      throw new Error('The journey day at completion is invalid.');
    }
    const title: string = taskTitle.trim();
    if (!title) {
      throw new Error('A task title is required to record completion.');
    }
    const completionDate: string = new Date().toISOString();

    const fields: Map<string, string> = await this._getFieldNames(COMPLETIONS_LIST_TITLE, [
      COMPLETION_JOURNEY_FIELD_TITLE,
      COMPLETION_TASK_FIELD_TITLE,
      COMPLETION_JOURNEY_DAY_FIELD_TITLE,
      COMPLETION_DATE_FIELD_TITLE
    ]);
    const journeyLookupIdField: string =
      `${this._getRequiredFieldName(fields, COMPLETION_JOURNEY_FIELD_TITLE)}Id`;
    const taskLookupIdField: string =
      `${this._getRequiredFieldName(fields, COMPLETION_TASK_FIELD_TITLE)}Id`;
    const journeyDayField: string = this._getRequiredFieldName(
      fields,
      COMPLETION_JOURNEY_DAY_FIELD_TITLE
    );
    const completionDateField: string = this._getRequiredFieldName(
      fields,
      COMPLETION_DATE_FIELD_TITLE
    );
    const item: Record<string, number | string> = {
      Title: title,
      [journeyLookupIdField]: journeyId,
      [taskLookupIdField]: taskId,
      [journeyDayField]: journeyDay,
      [completionDateField]: completionDate
    };
    const response = await this._spHttpClient.post(
      this._listItemsUrl(COMPLETIONS_LIST_TITLE),
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: SHAREPOINT_ACCEPT_HEADER,
          'Content-Type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify(item)
      }
    );

    if (!response.ok) {
      throw new Error(
        `Could not create a ${COMPLETIONS_LIST_TITLE} item (${response.status} ${response.statusText}).`
      );
    }

    const createdItem = await response.json() as ICreatedTaskCompletion;
    const completionId: unknown = createdItem.Id !== undefined ? createdItem.Id : createdItem.ID;
    if (typeof completionId !== 'number' || !Number.isInteger(completionId) || completionId < 1) {
      throw new Error(`The created ${COMPLETIONS_LIST_TITLE} item did not return a valid ID.`);
    }
    return completionId;
  }

  public async deleteCompletion(completionId: number): Promise<void> {
    this._validateId(completionId, 'completion');
    const response = await this._spHttpClient.post(
      `${this._listItemsUrl(COMPLETIONS_LIST_TITLE)}(${completionId})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: SHAREPOINT_ACCEPT_HEADER,
          'X-HTTP-Method': 'DELETE',
          'IF-MATCH': '*'
        }
      }
    );

    if (!response.ok) {
      throw new Error(
        `Could not delete the ${COMPLETIONS_LIST_TITLE} item (${response.status} ${response.statusText}).`
      );
    }
  }

  private async _getJourneyData(
    journeyId: number,
    includeContentType: boolean = false
  ): Promise<IOnboardingJourneyData> {
    if (!Number.isInteger(journeyId) || journeyId < 1) {
      throw new Error('The current SharePoint journey ID is invalid.');
    }

    const taskFieldTitles: string[] = [
      TASK_DAY_FIELD_TITLE,
      TASK_STATUS_FIELD_TITLE,
      TASK_DURATION_FIELD_TITLE,
      TASK_CLUSTER_FIELD_TITLE
    ];
    if (includeContentType) {
      taskFieldTitles.push(TASK_CONTENT_TYPE_FIELD_TITLE);
    }

    const [taskFields, completionFields] = await Promise.all([
      this._getFieldNames(TASKS_LIST_TITLE, taskFieldTitles),
      this._getFieldNames(COMPLETIONS_LIST_TITLE, [
        COMPLETION_JOURNEY_FIELD_TITLE,
        COMPLETION_TASK_FIELD_TITLE
      ])
    ]);

    const taskDayField: string = this._getRequiredFieldName(taskFields, TASK_DAY_FIELD_TITLE);
    const taskStatusField: string = this._getRequiredFieldName(taskFields, TASK_STATUS_FIELD_TITLE);
    const taskDurationField: string = this._getRequiredFieldName(taskFields, TASK_DURATION_FIELD_TITLE);
    const taskClusterField: string = this._getRequiredFieldName(taskFields, TASK_CLUSTER_FIELD_TITLE);
    const taskContentTypeField: string | undefined = includeContentType
      ? this._getRequiredFieldName(taskFields, TASK_CONTENT_TYPE_FIELD_TITLE)
      : undefined;
    const journeyLookupIdField: string =
      `${this._getRequiredFieldName(completionFields, COMPLETION_JOURNEY_FIELD_TITLE)}Id`;
    const taskLookupIdField: string =
      `${this._getRequiredFieldName(completionFields, COMPLETION_TASK_FIELD_TITLE)}Id`;

    const selectedTaskFields: string[] = [
      'Id',
      'Title',
      'Created',
      taskDayField,
      taskStatusField,
      taskDurationField,
      `${taskClusterField}/Title`
    ];
    if (taskContentTypeField) {
      selectedTaskFields.push(taskContentTypeField);
      selectedTaskFields.push(TASK_CONTENT_URL_FIELD_NAME);
    }
    const taskSelect: string = encodeURIComponent(selectedTaskFields.join(','));
    const tasksUrl: string = `${this._listItemsUrl(TASKS_LIST_TITLE)}?` +
      `$select=${taskSelect}&$expand=${encodeURIComponent(taskClusterField)}&$top=5000`;
    const completionSelect: string = encodeURIComponent(
      ['Id', journeyLookupIdField, taskLookupIdField].join(',')
    );
    const completionFilter: string = encodeURIComponent(`${journeyLookupIdField} eq ${journeyId}`);
    const completionsUrl: string = `${this._listItemsUrl(COMPLETIONS_LIST_TITLE)}?` +
      `$select=${completionSelect}&$filter=${completionFilter}&$top=5000`;

    const [tasks, completions] = await Promise.all([
      this._getAllItems<ITaskListItem>(tasksUrl, TASKS_LIST_TITLE),
      this._getAllItems<ITaskCompletionListItem>(completionsUrl, COMPLETIONS_LIST_TITLE)
    ]);

    return {
      taskDayField,
      taskStatusField,
      taskDurationField,
      taskClusterField,
      taskContentTypeField,
      taskLookupIdField,
      tasks,
      completions
    };
  }

  private _readCompletedTaskIds(
    completions: ITaskCompletionListItem[],
    taskLookupIdField: string
  ): Set<number> {
    const taskIds: Set<number> = new Set<number>();
    completions.forEach((completion: ITaskCompletionListItem) => {
      const taskId: unknown = completion[taskLookupIdField];
      if (typeof taskId !== 'number' || !Number.isInteger(taskId) || taskId < 1) {
        throw new Error(
          `A ${COMPLETIONS_LIST_TITLE} item for this journey has an invalid Task lookup.`
        );
      }
      taskIds.add(taskId);
    });
    return taskIds;
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

  private _validateId(value: number, entityName: string): void {
    if (!Number.isInteger(value) || value < 1) {
      throw new Error(`The ${entityName} ID is invalid.`);
    }
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
