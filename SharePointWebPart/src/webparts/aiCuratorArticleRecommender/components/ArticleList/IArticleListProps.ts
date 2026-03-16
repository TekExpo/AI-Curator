import { IArticle } from '../../services/TopicsService';

export interface IArticleListProps {
  articles: IArticle[];
  isLoading: boolean;
  errorMessage: string;
  infoMessage: string;
  /** Set of article URLs the user has already saved */
  savedArticleUrls: string[];
  /** Called when user clicks Save on an article card */
  onSaveArticle: (article: IArticle) => Promise<void>;
  /** Called when user clicks Share on an article card */
  onShareArticle: (article: IArticle) => void;
  /** Called when user clicks Share on LinkedIn on an article card */
  onLinkedInShareArticle: (article: IArticle) => void;
  /** Called when user clicks the Reload button to refresh recommendations */
  onReload: () => void;
}
