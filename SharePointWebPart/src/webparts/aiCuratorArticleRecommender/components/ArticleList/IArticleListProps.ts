import { IArticle } from '../../services/OpenAIService';

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
  /** Whether Viva Engage sharing is enabled */
  vivaEngageEnabled: boolean;
}
