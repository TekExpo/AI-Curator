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
  /** The URL of the LLM endpoint (set in config file) */
  llmEndpointUrl: string;
  /** Optional API key for direct OpenAI or Azure OpenAI calls (set in config file) */
  openAiApiKey: string;
  /** Optional model name for public OpenAI chat completions endpoints */
  openAiModel: string;
  /** System prompt used when calling an OpenAI-compatible endpoint */
  openAiSystemPrompt: string;
  /** Display name of the Articles list containing tags */
  articlesListName: string;
  /** Maximum number of articles to request from the LLM */
  maxArticles: number;
  /** Whether to cache LLM responses in sessionStorage */
  enableCaching: boolean;
  /** Display name of the userPersonalization list */
  userPersonalizationListName: string;
  /** Whether Viva Engage sharing is enabled */
  vivaEngageEnabled: boolean;
  /** Yammer App Client ID (only needed for Yammer REST API) */
  yammerClientId: string;
  /** Whether the current SharePoint theme is dark */
  isDarkTheme: boolean;
  /** Whether the web part is running inside Microsoft Teams */
  hasTeamsContext: boolean;
  /** The web part context (for pageContext, spHttpClient, msGraphClientFactory, etc.) */
  webPartContext: WebPartContext;
}
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
