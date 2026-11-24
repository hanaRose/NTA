import { readOnboardingTaskContentUrl } from './readOnboardingTaskContentUrl';

const baseUrl: string = 'https://contoso.sharepoint.com/sites/onboarding';

describe('readOnboardingTaskContentUrl', () => {
  it('reads an absolute URL string', () => {
    expect(readOnboardingTaskContentUrl('https://contoso.com/task', baseUrl))
      .toBe('https://contoso.com/task');
  });

  it('reads SharePoint hyperlink values and resolves relative URLs', () => {
    expect(readOnboardingTaskContentUrl({ Url: '/sites/onboarding/task' }, baseUrl))
      .toBe('https://contoso.sharepoint.com/sites/onboarding/task');
  });

  it('rejects empty or non-web URLs', () => {
    expect(() => readOnboardingTaskContentUrl('', baseUrl)).toThrow('ContentUrl');
    expect(() => readOnboardingTaskContentUrl('data:text/plain,task', baseUrl))
      .toThrow('HTTP or HTTPS');
  });
});
