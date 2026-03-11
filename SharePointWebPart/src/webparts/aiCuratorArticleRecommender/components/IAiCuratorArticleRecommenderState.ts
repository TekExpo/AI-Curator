import { IArticleRecommendation } from './IAiCuratorArticleRecommenderProps';

/**
 * Internal state managed by the root AiCuratorArticleRecommender component.
 */
export interface IAiCuratorArticleRecommenderState {
  /** Currently active Pivot tab key: 'tab1' | 'tab2' */
  activeTab: string;
  /** All unique tags extracted from the Articles list (Tab 1) */
  availableTags: string[];
  /** Tags currently selected by the user in Tab 1 */
  selectedTags: string[];
  /** Whether Tab 1 data is loading */
  isLoadingTags: boolean;
  /** Error message for Tab 1 */
  tab1Error: string;
  /** Success message for Tab 1 save operation */
  tab1Success: string;
  /** Articles returned from OpenAI (Tab 2) */
  articles: IArticleRecommendation[];
  /** Whether Tab 2 is loading recommendations */
  isLoadingArticles: boolean;
  /** Error message for Tab 2 */
  tab2Error: string;
  /** Info/warning message for Tab 2 (e.g. no interests saved) */
  tab2Info: string;
  /** The SharePanel is open for this article URL */
  sharePanelArticleUrl: string;
  /** The SharePanel is open for this article title */
  sharePanelArticleTitle: string;
  /** ID of the current user's userPersonalization list item (null if not yet created) */
  userPersonalizationItemId: number | null;
  /** Numeric SharePoint ID of the currently logged-in user */
  currentUserId: number | null;
  /** LoginName of the currently logged-in user */
  currentUserLoginName: string;
  /** SavedLinks string from the userPersonalization record */
  savedLinks: string;
}
