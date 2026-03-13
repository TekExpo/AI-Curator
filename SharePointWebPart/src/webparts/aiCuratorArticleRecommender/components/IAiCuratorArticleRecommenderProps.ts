import { WebPartContext } from '@microsoft/sp-webpart-base';

/**
 * Props passed from the web part class to the React component.
 */
export interface IAiCuratorArticleRecommenderProps {
  /** Display name of the Articles list (used for SP context only) */
  articlesListName: string;
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
