jest.mock('@microsoft/sp-http', () => ({
  SPHttpClient: {
    configurations: { v1: {} }
  }
}));

import type { SPHttpClient } from '@microsoft/sp-http';
import SharePointOnboardingTaskMetricsService from './SharePointOnboardingTaskMetricsService';

describe('SharePointOnboardingTaskMetricsService completion methods', () => {
  it('creates a completion with journey/task lookup IDs and supports undo', async () => {
    const fieldsResponse = {
      ok: true,
      status: 200,
      statusText: 'OK',
      json: jest.fn().mockResolvedValue({
        value: [
          { Title: 'Journey', EntityPropertyName: 'Journey' },
          { Title: 'Task', EntityPropertyName: 'Task' },
          {
            Title: 'JourneyDayAtComplete',
            EntityPropertyName: 'JourneyDayAtComplete'
          },
          { Title: 'CompletionDate', EntityPropertyName: 'CompletionDate' }
        ]
      })
    };
    const createdResponse = {
      ok: true,
      status: 201,
      statusText: 'Created',
      json: jest.fn().mockResolvedValue({ Id: 57 })
    };
    const deletedResponse = {
      ok: true,
      status: 204,
      statusText: 'No Content'
    };
    const spHttpClient = {
      get: jest.fn().mockResolvedValue(fieldsResponse),
      post: jest.fn()
        .mockResolvedValueOnce(createdResponse)
        .mockResolvedValueOnce(deletedResponse)
    } as unknown as SPHttpClient;
    const service = new SharePointOnboardingTaskMetricsService(
      spHttpClient,
      'https://contoso.sharepoint.com/sites/onboarding'
    );

    const createdAt: number = Date.now();
    await expect(service.createCompletion(12, 34, 'Welcome task', 8)).resolves.toBe(57);
    const createRequest = (spHttpClient.post as jest.Mock).mock.calls[0];
    expect(createRequest[0]).toBe(
      'https://contoso.sharepoint.com/sites/onboarding/_api/web/lists/GetByTitle(\'NTA_OB_TaskCompletions\')/items'
    );
    const payload = JSON.parse(createRequest[2].body);
    expect(payload).toEqual({
      Title: 'Welcome task',
      JourneyId: 12,
      TaskId: 34,
      JourneyDayAtComplete: 8,
      CompletionDate: expect.any(String)
    });
    expect(new Date(payload.CompletionDate).toISOString()).toBe(payload.CompletionDate);
    expect(new Date(payload.CompletionDate).getTime()).toBeGreaterThanOrEqual(createdAt);

    await expect(service.deleteCompletion(57)).resolves.toBeUndefined();
    const deleteRequest = (spHttpClient.post as jest.Mock).mock.calls[1];
    expect(deleteRequest[0]).toBe(
      'https://contoso.sharepoint.com/sites/onboarding/_api/web/lists/GetByTitle(\'NTA_OB_TaskCompletions\')/items(57)'
    );
    expect(deleteRequest[2].headers).toMatchObject({
      'X-HTTP-Method': 'DELETE',
      'IF-MATCH': '*'
    });
  });
});
