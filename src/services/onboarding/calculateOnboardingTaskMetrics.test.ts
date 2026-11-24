import {
  calculateOnboardingTaskMetrics,
  ICompletedTaskReference,
  ITaskMetricCandidate
} from './calculateOnboardingTaskMetrics';

describe('calculateOnboardingTaskMetrics', () => {
  it('counts published tasks and sums duration for due tasks without a completion', () => {
    const tasks: ITaskMetricCandidate[] = [
      { id: 1, isPublished: true, clusterName: 'תחנת מוצא', journeyDay: 1, duration: 12 },
      { id: 2, isPublished: true, clusterName: 'תחנת מוצא', journeyDay: 7, duration: 5 },
      { id: 3, isPublished: true, clusterName: 'תחנת מוצא', journeyDay: 8, duration: 9 },
      { id: 4, isPublished: false, clusterName: 'מרכז בקרה', journeyDay: 2, duration: 20 }
    ];
    const completions: ICompletedTaskReference[] = [
      { taskId: 1 },
      { taskId: 1 },
      { taskId: 4 }
    ];

    const metrics = calculateOnboardingTaskMetrics(tasks, completions, 7);

    expect(metrics).toMatchObject({
      completedTasks: 1,
      totalTasks: 3,
      openTaskCount: 1,
      learningMinutes: 5
    });
    expect(metrics.clusterProgress.find((cluster) => cluster.name === 'תחנת מוצא')).toEqual({
      name: 'תחנת מוצא',
      totalTasks: 3,
      completedTasks: 1
    });
  });
});
