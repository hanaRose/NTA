import {
  IOnboardingClusterMetrics,
  IOnboardingTaskMetrics,
  ONBOARDING_CLUSTER_NAMES
} from './IOnboardingTaskMetrics';

export interface ITaskMetricCandidate {
  id: number;
  isPublished: boolean;
  clusterName: string;
  journeyDay?: number;
  duration?: number;
}

export interface ICompletedTaskReference {
  taskId: number;
}

export function calculateOnboardingTaskMetrics(
  tasks: ITaskMetricCandidate[],
  completions: ICompletedTaskReference[],
  currentDay: number
): IOnboardingTaskMetrics {
  const publishedTasks: ITaskMetricCandidate[] = tasks.filter(
    (task: ITaskMetricCandidate) => task.isPublished
  );
  const publishedTaskIds: Set<number> = new Set<number>(
    publishedTasks.map((task: ITaskMetricCandidate) => task.id)
  );
  const completedTaskIds: Set<number> = new Set<number>(
    completions
      .map((completion: ICompletedTaskReference) => completion.taskId)
      .filter((taskId: number) => publishedTaskIds.has(taskId))
  );
  const openTasks: ITaskMetricCandidate[] = publishedTasks.filter(
    (task: ITaskMetricCandidate) =>
      task.journeyDay !== undefined &&
      task.journeyDay <= currentDay &&
      !completedTaskIds.has(task.id)
  );
  const extraClusterNames: string[] = [];
  publishedTasks.forEach((task: ITaskMetricCandidate) => {
    if (ONBOARDING_CLUSTER_NAMES.indexOf(task.clusterName) === -1 &&
      extraClusterNames.indexOf(task.clusterName) === -1) {
      extraClusterNames.push(task.clusterName);
    }
  });
  const clusterNames: string[] = [...ONBOARDING_CLUSTER_NAMES, ...extraClusterNames];
  const clusterProgress: IOnboardingClusterMetrics[] = clusterNames.map(
    (name: string) => {
      const clusterTasks: ITaskMetricCandidate[] = publishedTasks.filter(
        (task: ITaskMetricCandidate) => task.clusterName === name
      );
      return {
        name,
        totalTasks: clusterTasks.length,
        completedTasks: clusterTasks.filter(
          (task: ITaskMetricCandidate) => completedTaskIds.has(task.id)
        ).length
      };
    }
  );

  openTasks.forEach((task: ITaskMetricCandidate) => {
    if (task.duration !== undefined && task.duration < 0) {
      throw new Error('An open task cannot have a negative duration.');
    }
  });

  return {
    completedTasks: completedTaskIds.size,
    totalTasks: publishedTasks.length,
    openTaskCount: openTasks.length,
    learningMinutes: openTasks.reduce(
      (total: number, task: ITaskMetricCandidate) => total + (task.duration || 0),
      0
    ),
    clusterProgress
  };
}
