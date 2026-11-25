import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { IOnboardingTip } from '../../services/tips/IOnboardingTip';
import { IOnboardingTipsService } from '../../services/tips/IOnboardingTipsService';
import SharePointOnboardingTipsService from '../../services/tips/SharePointOnboardingTipsService';
import TipsBar from './components/TipsBar';
import { ITipsBarProps } from './components/ITipsBarProps';

export default class TipsBarWebPart extends BaseClientSideWebPart<Record<string, never>> {
  private _tipsService: IOnboardingTipsService | undefined;
  private _tips: IOnboardingTip[] = [];
  private _isLoading: boolean = true;
  private _loadError: string | undefined;
  private _isTipsBarDisposed: boolean = false;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._tipsService = new SharePointOnboardingTipsService(
      this.context.spHttpClient,
      this.context.pageContext.web.absoluteUrl
    );
    this.loadTips().catch((error: unknown) => {
      console.error('Unexpected error while loading onboarding tips.', error);
    });
  }

  public render(): void {
    const element: React.ReactElement<ITipsBarProps> = React.createElement(
      TipsBar,
      {
        tips: this._tips,
        isLoading: this._isLoading,
        loadError: this._loadError
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    this._isTipsBarDisposed = true;
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private async loadTips(): Promise<void> {
    this._loadError = undefined;
    try {
      if (!this._tipsService) {
        throw new Error('The onboarding tips service is not available.');
      }
      this._tips = await this._tipsService.getPublishedTips();
    } catch (error) {
      this._loadError = error instanceof Error ? error.message : String(error);
      console.error('Could not load onboarding tips.', error);
    } finally {
      this._isLoading = false;
      if (!this._isTipsBarDisposed) {
        this.render();
      }
    }
  }
}
