import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { IOnboardingProgressService } from '../../services/onboarding/IOnboardingProgressService';
import { IOnboardingTask } from '../../services/onboarding/IOnboardingTask';
import { IOnboardingTasksService } from '../../services/onboarding/IOnboardingTasksService';
import {
  ONBOARDING_TASK_COMPLETIONS_CHANGED,
  notifyOnboardingTaskCompletionsChanged
} from '../../services/onboarding/onboardingTaskEvents';
import SharePointOnboardingProgressService from '../../services/onboarding/SharePointOnboardingProgressService';
import SharePointOnboardingTaskMetricsService from '../../services/onboarding/SharePointOnboardingTaskMetricsService';
import TasksSection from './components/TasksSection';
import { ITasksSectionProps } from './components/ITasksSectionProps';

export default class TasksSectionWebPart extends BaseClientSideWebPart<Record<string, never>> {
  private _progressService: IOnboardingProgressService | undefined;
  private _tasksService: IOnboardingTasksService | undefined;
  private _currentDay: number = 1;
  private _journeyId: number | undefined;
  private _tasks: IOnboardingTask[] = [];
  private _isLoading: boolean = true;
  private _loadError: string | undefined;
  private _isTasksSectionDisposed: boolean = false;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._progressService = new SharePointOnboardingProgressService(
      this.context.spHttpClient,
      this.context.pageContext.web.absoluteUrl
    );
    this._tasksService = new SharePointOnboardingTaskMetricsService(
      this.context.spHttpClient,
      this.context.pageContext.web.absoluteUrl
    );
    window.addEventListener(
      ONBOARDING_TASK_COMPLETIONS_CHANGED,
      this._onTaskCompletionsChanged
    );
    this.loadTasks().catch((error: unknown) => {
      console.error('Unexpected error while loading onboarding tasks.', error);
    });
  }

  public render(): void {
    const element: React.ReactElement<ITasksSectionProps> = React.createElement(
      TasksSection,
      {
        currentDay: this._currentDay,
        tasks: this._tasks,
        isLoading: this._isLoading,
        loadError: this._loadError,
        onCompleteTask: this._completeTask,
        onUndoCompletion: this._undoCompletion
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    this._isTasksSectionDisposed = true;
    window.removeEventListener(
      ONBOARDING_TASK_COMPLETIONS_CHANGED,
      this._onTaskCompletionsChanged
    );
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private async loadTasks(): Promise<void> {
    this._loadError = undefined;
    try {
      if (!this._progressService || !this._tasksService) {
        throw new Error('The onboarding task services are not available.');
      }
      const progress = await this._progressService.getForUser(this.context.pageContext.user.email);
      this._journeyId = progress.journeyId;
      this._currentDay = progress.currentDay;
      this._tasks = await this._tasksService.getIncompleteForJourney(progress.journeyId);
    } catch (error) {
      this._loadError = error instanceof Error ? error.message : String(error);
      console.error('Could not load onboarding tasks.', error);
    } finally {
      this._isLoading = false;
      if (!this._isTasksSectionDisposed) {
        this.render();
      }
    }
  }

  private _onTaskCompletionsChanged = (): void => {
    if (this._isTasksSectionDisposed) {
      return;
    }
    this.loadTasks().catch((error: unknown) => {
      console.error('Unexpected error while refreshing onboarding tasks.', error);
    });
  };

  private _completeTask = async (task: IOnboardingTask): Promise<number> => {
    if (!this._tasksService || this._journeyId === undefined) {
      throw new Error('The current onboarding journey is not available.');
    }
    const completionId: number = await this._tasksService.createCompletion(
      this._journeyId,
      task.id,
      task.title,
      this._currentDay
    );
    notifyOnboardingTaskCompletionsChanged();
    return completionId;
  };

  private _undoCompletion = async (completionId: number): Promise<void> => {
    if (!this._tasksService) {
      throw new Error('The onboarding task service is not available.');
    }
    await this._tasksService.deleteCompletion(completionId);
    notifyOnboardingTaskCompletionsChanged();
  };
}
