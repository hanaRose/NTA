import { getOnboardingTaskButtonLabel } from './getOnboardingTaskButtonLabel';

describe('getOnboardingTaskButtonLabel', () => {
  it.each([
    ['סרטון', 'צפיה בסרטון'],
    ['מסמך אניטרקטיבי', 'לפתיחת המסמך'],
    ['לומדה', 'להתחלת הלומדה'],
    ['מפגש', 'להתחברות למפגש'],
    ['מיניסייט', 'לפתיחת האתר'],
    ['טופס', 'למילוי הטופס']
  ])('maps %s to its action label', (taskContentType, expectedLabel) => {
    expect(getOnboardingTaskButtonLabel(taskContentType)).toBe(expectedLabel);
  });

  it('uses the document action for the established spelling variant', () => {
    expect(getOnboardingTaskButtonLabel('מסמך אינטרקטיבי')).toBe('לפתיחת המסמך');
  });

  it('uses a generic label for an unrecognized content type', () => {
    expect(getOnboardingTaskButtonLabel('סוג חדש')).toBe('לפתיחת המשימה');
  });
});
