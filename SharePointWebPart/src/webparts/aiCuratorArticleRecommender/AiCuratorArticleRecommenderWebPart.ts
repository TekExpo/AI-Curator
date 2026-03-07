/**
 * ============================================================================
 * AI Curator - Article Recommender Web Part
 * ============================================================================
 *
 * *** HEFT-BASED BUILD (SPFx v1.22.1+) ***
 * Migrated from legacy Gulp toolchain to modern Heft build rig.
 * Build rig: @microsoft/spfx-web-build-rig
 *
 * SPFx Web Part entry point. This class:
 *  - Initializes PnPjs with the SPFx context
 *  - Configures the Property Pane with all user-facing settings
 *  - Renders the React component with the configured properties
 *
 * DEPLOYMENT (Heft-based):
 *  1. npm install
 *  2. npm run build              (heft build --clean)
 *  3. npm run bundle             (heft bundle --production)
 *  4. npm run package-solution   (heft package-solution --production)
 *  5. Upload the .sppkg from sharepoint/solution/ to your App Catalog
 *  6. Add the web part to a SharePoint page
 *
 * EXTENDING THE LLM PAYLOAD:
 *  If your LLM endpoint requires additional fields (e.g., userId, language,
 *  category filters), add new Property Pane fields below and pass them through
 *  to the React component via IAiCuratorArticleRecommenderProps. Then update
 *  the fetch call in AiCuratorArticleRecommender.tsx to include those fields
 *  in the POST body.
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

import { spfi, SPFx } from '@pnp/sp';
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';

import AiCuratorArticleRecommender from './components/AiCuratorArticleRecommender';
import { aiCuratorArticleRecommenderConfig } from './AiCuratorArticleRecommender.config';
import { IAiCuratorArticleRecommenderProps } from './components/IAiCuratorArticleRecommenderProps';
import * as strings from 'AiCuratorArticleRecommenderWebPartStrings';

/**
 * Properties exposed via the Property Pane for this web part.
 * All fields are persisted in the web part's property bag on the page.
 */
export interface IAiCuratorArticleRecommenderWebPartProps {
  /** Optional model name for public OpenAI chat completions endpoints */
  openAiModel: string;
  /** System prompt used when calling an OpenAI-compatible endpoint */
  openAiSystemPrompt: string;
  /** The SharePoint list name containing keyword data */
  listName: string;
  /** The internal column name that holds keyword values */
  keywordColumnName: string;
  /** Optional: target site URL. If empty, uses the current site context */
  siteUrl: string;
  /** Maximum number of articles the LLM should return */
  maxArticles: number;
  /** Whether to cache LLM responses in sessionStorage */
  enableCaching: boolean;
}

export default class AiCuratorArticleRecommenderWebPart extends BaseClientSideWebPart<IAiCuratorArticleRecommenderWebPartProps> {

  private _sp!: SPFI;
  private _isDarkTheme: boolean = false;

  /**
   * Called once when the web part is first loaded.
   * Initializes PnPjs with the SPFx context so all subsequent SP REST
   * calls are properly authenticated.
   */
  protected async onInit(): Promise<void> {
    await super.onInit();

    // Initialize PnPjs with the SPFx context for authenticated SharePoint calls
    this._sp = spfi().using(SPFx(this.context));

    return;
  }

  /**
   * Renders the React component into the web part's DOM element.
   * Called on initial load and whenever a property pane value changes.
   */
  public render(): void {
    const element: React.ReactElement<IAiCuratorArticleRecommenderProps> = React.createElement(
      AiCuratorArticleRecommender,
      {
        llmEndpointUrl: aiCuratorArticleRecommenderConfig.llmEndpointUrl,
        openAiApiKey: aiCuratorArticleRecommenderConfig.openAiApiKey,
        openAiModel: this.properties.openAiModel,
        openAiSystemPrompt: this.properties.openAiSystemPrompt,
        listName: this.properties.listName,
        keywordColumnName: this.properties.keywordColumnName,
        siteUrl: this.properties.siteUrl,
        maxArticles: this.properties.maxArticles,
        enableCaching: this.properties.enableCaching,
        spInstance: this._sp,
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks?.microsoftTeams,
        webPartContext: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  /**
   * Cleans up the React component when the web part is disposed.
   * Prevents memory leaks by unmounting the React tree.
   */
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * Ensures the Property Pane re-renders the component on every change
   * (reactive mode) so users see live updates.
   */
  protected get disableReactivePropertyChanges(): boolean {
    return false;
  }

  /**
   * Defines the Property Pane layout with all configurable fields.
   *
   * GROUP 1 – LLM Configuration:
  *   - OpenAI Model
  *   - OpenAI System Prompt
   *   - Max Articles slider
   *   - Enable Caching toggle
   *
   * GROUP 2 – SharePoint Data Source:
   *   - List Name
   *   - Keyword Column Name
   *   - Site URL (optional override)
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.LlmConfigGroupName,
              groupFields: [
                PropertyPaneTextField('openAiModel', {
                  label: strings.OpenAiModelFieldLabel,
                  placeholder: 'gpt-4.1-mini',
                  description: 'Required for public OpenAI endpoints. Leave empty for Azure OpenAI deployment URLs that already target a deployment.',
                  multiline: false
                }),
                PropertyPaneTextField('openAiSystemPrompt', {
                  label: strings.OpenAiSystemPromptFieldLabel,
                  placeholder: 'Return only a JSON array of article recommendations for the provided keywords.',
                  description: 'System prompt sent to OpenAI-compatible endpoints. The model must return JSON only.',
                  multiline: true
                }),
                PropertyPaneSlider('maxArticles', {
                  label: strings.MaxArticlesFieldLabel,
                  min: 1,
                  max: 25,
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
            {
              groupName: strings.DataSourceGroupName,
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel,
                  placeholder: 'Articles',
                  description: 'The display name of the SharePoint list containing keywords.'
                }),
                PropertyPaneTextField('keywordColumnName', {
                  label: strings.KeywordColumnNameFieldLabel,
                  placeholder: 'Keywords',
                  description: 'The internal name of the column that stores keywords.'
                }),
                PropertyPaneTextField('siteUrl', {
                  label: strings.SiteUrlFieldLabel,
                  placeholder: 'https://tenant.sharepoint.com/sites/MySite',
                  description: 'Optional. Leave empty to use the current site.'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
