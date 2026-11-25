jest.mock('@microsoft/sp-http', () => ({
  SPHttpClient: {
    configurations: { v1: {} }
  }
}));

import type { SPHttpClient } from '@microsoft/sp-http';
import SharePointOnboardingTipsService from './SharePointOnboardingTipsService';

describe('SharePointOnboardingTipsService', () => {
  it('returns published tips newest first with lookup, link, and QR values mapped', async () => {
    const fieldsResponse = {
      ok: true,
      status: 200,
      statusText: 'OK',
      json: jest.fn().mockResolvedValue({
        value: [
          { Title: 'אשכול', EntityPropertyName: 'Streak' },
          { Title: 'יום במסע', EntityPropertyName: 'JourneyDay' },
          { Title: 'מצב פרסום', EntityPropertyName: 'PublicationStatus' },
          { Title: 'תוכן הטיפ', EntityPropertyName: 'TipText' },
          { Title: 'קישור', EntityPropertyName: 'TipLink' },
          { Title: 'מלל הכפתור', EntityPropertyName: 'ButtonLabel' },
          { Title: 'הצגת QR', EntityPropertyName: 'ShowQR' },
          { Title: 'תמונה', EntityPropertyName: 'TipImage' }
        ]
      })
    };
    const tipsResponse = {
      ok: true,
      status: 200,
      statusText: 'OK',
      json: jest.fn().mockResolvedValue({
        value: [
          {
            Id: 2,
            Title: 'Newest tip',
            Created: '2026-11-24T10:00:00Z',
            Streak: { Title: 'People' },
            JourneyDay: 4,
            PublicationStatus: 'פורסם',
            TipText: 'Tip body',
            TipLink: { Url: '/sites/onboarding/SitePages/tip.aspx' },
            ButtonLabel: 'לפתיחת המידע',
            ShowQR: true,
            TipImage: JSON.stringify({
              type: 'thumbnail',
              nativeFile: {
                serverUrl: 'https://contoso.sharepoint.com',
                serverRelativeUrl: '/sites/onboarding/SiteAssets/tip-qr.png'
              }
            })
          },
          {
            Id: 1,
            Title: 'Older tip',
            Created: '2026-11-23T10:00:00Z',
            Streak: { Title: 'People' },
            JourneyDay: 2,
            PublicationStatus: 'פורסם',
            TipText: 'Older body',
            TipLink: null,
            ButtonLabel: null,
            ShowQR: false,
            TipImage: null
          },
          {
            Id: 3,
            Title: 'Draft tip',
            Created: '2026-11-25T10:00:00Z',
            Streak: { Title: 'People' },
            JourneyDay: 5,
            PublicationStatus: 'טיוטה',
            TipText: 'Not published',
            TipLink: null,
            ButtonLabel: null,
            ShowQR: false,
            TipImage: null
          }
        ]
      })
    };
    const spHttpClient = {
      get: jest.fn()
        .mockResolvedValueOnce(fieldsResponse)
        .mockResolvedValueOnce(tipsResponse)
    } as unknown as SPHttpClient;
    const service = new SharePointOnboardingTipsService(
      spHttpClient,
      'https://contoso.sharepoint.com/sites/onboarding'
    );

    const tips = await service.getPublishedTips();

    expect(tips).toHaveLength(2);
    expect(tips[0]).toEqual({
      id: 2,
      title: 'Newest tip',
      text: 'Tip body',
      day: 4,
      category: 'People',
      link: 'https://contoso.sharepoint.com/sites/onboarding/SitePages/tip.aspx',
      linkLabel: 'לפתיחת המידע',
      showQr: true,
      qrImageUrl: 'https://contoso.sharepoint.com/sites/onboarding/SiteAssets/tip-qr.png'
    });
    expect(tips[1].title).toBe('Older tip');

    const itemsUrl = (spHttpClient.get as jest.Mock).mock.calls[1][0] as string;
    expect(decodeURIComponent(itemsUrl)).toContain(
      "$filter=PublicationStatus eq 'פורסם'"
    );
    expect(decodeURIComponent(itemsUrl)).toContain('$orderby=Created desc');
  });
});
