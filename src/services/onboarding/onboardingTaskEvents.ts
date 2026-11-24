export const ONBOARDING_TASK_COMPLETIONS_CHANGED: string =
  'nta:onboarding-task-completions-changed';

export function notifyOnboardingTaskCompletionsChanged(): void {
  window.dispatchEvent(new Event(ONBOARDING_TASK_COMPLETIONS_CHANGED));
}
