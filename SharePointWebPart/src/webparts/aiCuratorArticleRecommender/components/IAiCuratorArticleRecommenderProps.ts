import { SPFI } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';

/**
 * Represents a single article recommendation returned by the LLM endpoint.
 */
export interface IArticleRecommendation {
  /** The article title (displayed as link text) */
  title: string;
  /** The full URL to the article (opens in a new tab) */
  url: string;
  /** A short summary or description of the article */
  summary: string;
}

/**
 * Props passed from the web part class to the React component.
 */
export interface IAiCuratorArticleRecommenderProps {
  /** The URL of the LLM endpoint */
  llmEndpointUrl: string;
  /** SharePoint list name containing keyword data */
  listName: string;
  /** Internal column name holding keyword values */
  keywordColumnName: string;
  /** Optional site URL override; empty string = current site */
  siteUrl: string;
  /** Maximum number of articles to request from the LLM */
  maxArticles: number;
  /** Whether to cache LLM responses in sessionStorage */
  enableCaching: boolean;
  /** PnPjs SPFI instance initialized with SPFx context */
  spInstance: SPFI;
  /** Whether the current SharePoint theme is dark */
  isDarkTheme: boolean;
  /** Whether the web part is running inside Microsoft Teams */
  hasTeamsContext: boolean;
  /** The web part context (for pageContext, httpClient, etc.) */
  webPartContext: WebPartContext;
}

/**
 * Internal state managed by the React component via useState hooks.
 */
export interface IAiCuratorArticleRecommenderState {
  /** The list of article recommendations returned by the LLM */
  articles: IArticleRecommendation[];
  /** Whether a fetch operation is in progress */
  isLoading: boolean;
  /** Error message to display (empty string = no error) */
  errorMessage: string;
  /** The raw keywords extracted from SharePoint */
  keywords: string;
}
