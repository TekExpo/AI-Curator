/**
 * ============================================================================
 * AI Curator - Article Recommender Web Part
 * ============================================================================
 *
 * *** HEFT-BASED BUILD (SPFx v1.22.1+) ***
 * Build rig: @microsoft/spfx-web-build-rig
 *
 * DEPLOYMENT (Heft-based):
 *  1. npm install
 *  2. npm run build              (heft build --clean)
 *  3. npm run bundle             (heft bundle --production)
 *  4. npm run package-solution   (heft package-solution --production)
 *  5. Upload the .sppkg from sharepoint/solution/ to your App Catalog
 *
 * TOPICS DATA FLOW:
 *  Topics are discovered via the external suggest-topics API.
 *  Article recommendations come from the external articles API.
 *  No OpenAI keys or endpoints are stored in this web part.
 *
 * ============================================================================
 */

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import AiCuratorArticleRecommender from './components/AiCuratorArticleRecommender';
import { IAiCuratorArticleRecommenderProps } from './components/IAiCuratorArticleRecommenderProps';
import { SharePointService } from './services/SharePointService';
import * as strings from 'AiCuratorArticleRecommenderWebPartStrings';

/**
 * Properties persisted in the web part property bag.
 * Keywords are NEVER stored here — they come from userPersonalization at runtime.
 */
export interface IAiCuratorArticleRecommenderWebPartProps {
  // ── SharePoint Data Source ────────────────────────────────────────
  articlesListName: string;
  // ── Personalization & Sharing ─────────────────────────────────────
  userPersonalizationListName: string;
  articlesLimit: number;
}

export default class AiCuratorArticleRecommenderWebPart extends BaseClientSideWebPart<IAiCuratorArticleRecommenderWebPartProps> {

  private _isDarkTheme: boolean = false;

  protected async onInit(): Promise<void> {
    await super.onInit();
    try {
      const spService = new SharePointService(this.context);
      await spService.ensureSiteLists(
        this.properties.articlesListName || 'Articles',
        this.properties.userPersonalizationListName || 'UserPersonalization'
      );
    } catch (err) {
      // Non-fatal: log and continue — lists may already exist or user may lack permissions
      console.warn('AI Curator: ensureSiteLists failed', err);
    }
  }

  public render(): void {
    const element: React.ReactElement<IAiCuratorArticleRecommenderProps> = React.createElement(
      AiCuratorArticleRecommender,
      {
        articlesListName: this.properties.articlesListName || 'Articles',
        userPersonalizationListName: this.properties.userPersonalizationListName || 'userPersonalization',
        articlesLimit: this.properties.articlesLimit ?? 10,
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks?.microsoftTeams,
        webPartContext: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get disableReactivePropertyChanges(): boolean {
    return false;
  }

  /**
   * Property pane — two groups.
   * NO keyword, tag, or OpenAI fields anywhere in this pane.
   * Topics are discovered via external API at runtime.
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            // ── Group 1: SharePoint Data Source ───────────────────────────
            {
              groupName: strings.DataSourceGroupName,
              groupFields: [
                PropertyPaneTextField('articlesListName', {
                  label: strings.ListNameFieldLabel,
                  placeholder: 'Articles',
                  description: 'Display name of the Articles list (used for context only).',
                  value: 'Articles'
                })
              ]
            },
            // ── Group 2: Personalization & Sharing ────────────────────────
            {
              groupName: strings.PersonalizationGroupName,
              groupFields: [
                PropertyPaneTextField('userPersonalizationListName', {
                  label: strings.UserPersonalizationListNameFieldLabel,
                  description: 'Display name of the list storing user tag selections and saved links.',
                  value: 'userPersonalization'
                }),
                PropertyPaneSlider('articlesLimit', {
                  label: strings.ArticlesLimitFieldLabel,
                  min: 1,
                  max: 100,
                  step: 1,
                  showValue: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
