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
 * KEYWORDS DATA FLOW:
 *  Keywords for OpenAI are NEVER read from the property pane.
 *  They come exclusively from userPersonalization.SelectedTags
 *  matched by the logged-in user's numeric SharePoint ID.
 *
 * ============================================================================
 */

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import AiCuratorArticleRecommender from './components/AiCuratorArticleRecommender';
import { IAiCuratorArticleRecommenderProps } from './components/IAiCuratorArticleRecommenderProps';
import * as strings from 'AiCuratorArticleRecommenderWebPartStrings';

/**
 * Properties persisted in the web part property bag.
 * Keywords are NEVER stored here — they come from userPersonalization at runtime.
 */
export interface IAiCuratorArticleRecommenderWebPartProps {
  // ── LLM Configuration ────────────────────────────────────────────
  openAiModel: string;
  openAiSystemPrompt: string;
  maxArticles: number;
  enableCaching: boolean;
  // ── SharePoint Data Source ────────────────────────────────────────
  articlesListName: string;
  // ── Personalization & Sharing ─────────────────────────────────────
  userPersonalizationListName: string;
  vivaEngageEnabled: boolean;
  yammerClientId: string;
}

export default class AiCuratorArticleRecommenderWebPart extends BaseClientSideWebPart<IAiCuratorArticleRecommenderWebPartProps> {

  private _isDarkTheme: boolean = false;

  protected async onInit(): Promise<void> {
    await super.onInit();
    return;
  }

  public render(): void {
    const element: React.ReactElement<IAiCuratorArticleRecommenderProps> = React.createElement(
      AiCuratorArticleRecommender,
      {
        openAiModel: this.properties.openAiModel,
        openAiSystemPrompt: this.properties.openAiSystemPrompt,
        articlesListName: this.properties.articlesListName || 'Articles',
        maxArticles: this.properties.maxArticles || 5,
        enableCaching: this.properties.enableCaching !== false,
        userPersonalizationListName: this.properties.userPersonalizationListName || 'userPersonalization',
        vivaEngageEnabled: !!this.properties.vivaEngageEnabled,
        yammerClientId: this.properties.yammerClientId || '',
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
   * Final property pane — three groups as specified.
   * NO keyword or tag input fields anywhere in this pane.
   * Keywords come exclusively from userPersonalization.SelectedTags at runtime.
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            // ── Group 1: LLM Configuration ────────────────────────────────
            {
              groupName: strings.LlmConfigGroupName,
              groupFields: [
                PropertyPaneTextField('openAiModel', {
                  label: strings.OpenAiModelFieldLabel,
                  placeholder: 'gpt-4o',
                  description: 'Model name for public OpenAI endpoints (e.g. gpt-4o). Leave empty for Azure OpenAI deployment URLs.',
                  multiline: false
                }),
                PropertyPaneTextField('openAiSystemPrompt', {
                  label: strings.OpenAiSystemPromptFieldLabel,
                  placeholder: 'Return only a JSON array of article recommendations for the provided keywords.',
                  description: 'System prompt sent to the OpenAI endpoint. Must instruct the model to return JSON only.',
                  multiline: true
                }),
                PropertyPaneSlider('maxArticles', {
                  label: strings.MaxArticlesFieldLabel,
                  min: 1,
                  max: 10,
                  step: 1,
                  showValue: true,
                  value: 5
                }),
                PropertyPaneToggle('enableCaching', {
                  label: strings.EnableCachingFieldLabel,
                  onText: 'Caching enabled',
                  offText: 'Caching disabled',
                  checked: true
                })
              ]
            },
            // ── Group 2: SharePoint Data Source ───────────────────────────
            {
              groupName: strings.DataSourceGroupName,
              groupFields: [
                PropertyPaneTextField('articlesListName', {
                  label: strings.ListNameFieldLabel,
                  placeholder: 'Articles',
                  description: 'Display name of the Articles list containing tags (Keywords column). Keywords are shown on the My Interests tab — they are never entered via the property pane.',
                  value: 'Articles'
                })
                // ← No keyword column name field
                // ← No keyword input field
                // Keywords are read at runtime from userPersonalization.SelectedTags only
              ]
            },
            // ── Group 3: Personalization & Sharing ────────────────────────
            {
              groupName: strings.PersonalizationGroupName,
              groupFields: [
                PropertyPaneTextField('userPersonalizationListName', {
                  label: strings.UserPersonalizationListNameFieldLabel,
                  description: 'Display name of the list storing user tag selections and saved links.',
                  value: 'userPersonalization'
                }),
                PropertyPaneToggle('vivaEngageEnabled', {
                  label: strings.VivaEngageEnabledFieldLabel,
                  onText: 'Enabled',
                  offText: 'Disabled'
                }),
                PropertyPaneTextField('yammerClientId', {
                  label: strings.YammerClientIdFieldLabel,
                  description: 'Required only if using Yammer REST API instead of Microsoft Graph.'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
