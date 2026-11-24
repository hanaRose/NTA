const TASK_BUTTON_LABELS: Readonly<Record<string, string>> = {
  'סרטון': 'צפיה בסרטון',
  'מסמך אניטרקטיבי': 'לפתיחת המסמך',
  'מסמך אינטרקטיבי': 'לפתיחת המסמך',
  'מסמך אינטראקטיבי': 'לפתיחת המסמך',
  'לומדה': 'להתחלת הלומדה',
  'מפגש': 'להתחברות למפגש',
  'מיניסייט': 'לפתיחת האתר',
  'טופס': 'למילוי הטופס'
};

export function getOnboardingTaskButtonLabel(taskContentType: string): string {
  const normalizedType: string = taskContentType.trim().replace(/\s+/g, ' ');
  return TASK_BUTTON_LABELS[normalizedType] || 'לפתיחת המשימה';
}
