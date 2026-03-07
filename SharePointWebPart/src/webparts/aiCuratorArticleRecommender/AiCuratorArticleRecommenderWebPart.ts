/**
 * ============================================================================
 * AI Curator - Article Recommender Web Part
 * ============================================================================
 *
 * SPFx Web Part entry point. This class:
 *  - Initializes PnPjs with the SPFx context
 *  - Configures the Property Pane with all user-facing settings
 *  - Renders the React component with the configured properties
 *
 * DEPLOYMENT:
 *  1. npm install
 *  2. gulp build
 *  3. gulp bundle --ship
 *  4. gulp package-solution --ship
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
import { IAiCuratorArticleRecommenderProps } from './components/IAiCuratorArticleRecommenderProps';
import * as strings from 'AiCuratorArticleRecommenderWebPartStrings';

/**
 * Properties exposed via the Property Pane for this web part.
 * All fields are persisted in the web part's property bag on the page.
 */
export interface IAiCuratorArticleRecommenderWebPartProps {
  /** The URL of the LLM endpoint that returns article recommendations */
  llmEndpointUrl: string;
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

  private _sp: SPFI;
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
        llmEndpointUrl: this.properties.llmEndpointUrl,
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
   *   - LLM Endpoint URL (required)
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
                PropertyPaneTextField('llmEndpointUrl', {
                  label: strings.LlmEndpointUrlFieldLabel,
                  placeholder: 'https://your-llm-api.example.com/recommend',
                  description: 'The full URL of the LLM endpoint that accepts keyword-based article requests.',
                  multiline: false
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
