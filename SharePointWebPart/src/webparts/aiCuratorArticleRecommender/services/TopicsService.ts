export interface IArticle {
  title: string;
  summary?: string;
  url: string;
  source?: string;
  published?: string;
  topic?: string;
  tags?: string;
}

const BASE_URL = 'https://ai-curator.azurewebsites.net';

/**
 * Service for the external AI Curator API endpoints.
 * Uses native fetch() — no SPHttpClient or PnPjs dependency.
 */
export class TopicsService {
  /**
   * Fetches suggested topics from the suggest-topics endpoint.
   * @param query The user's search text
   * @returns Array of topic strings, or [] if none / missing from response
   */
  public async getSuggestedTopics(query: string): Promise<string[]> {
    const url = `${BASE_URL}/suggest-topics?query=${encodeURIComponent(query)}`;
    let response: Response;
    try {
      response = await fetch(url);
    } catch (networkErr) {
      throw new Error(`Network error fetching topics: ${String(networkErr)}`);
    }
    if (!response.ok) {
      throw new Error(`Topics API returned HTTP ${response.status}`);
    }
    const data = await response.json() as { topics?: unknown };
    if (!Array.isArray(data.topics)) return [];
    return (data.topics as unknown[]).filter(
      (t): t is string => typeof t === 'string' && t.length > 0
    );
  }

  /**
   * Fetches article recommendations from the articles endpoint.
   * @param selectedTags Comma-separated topic string from userPersonalization.SelectedTags
   * @param limit Maximum number of articles to return (default: 20)
   * @returns Array of IArticle objects
   */
  public async getArticles(selectedTags: string, limit: number = 20): Promise<IArticle[]> {
    const url =
      `${BASE_URL}/articles` +
      `?topics=${encodeURIComponent(selectedTags)}&limit=${limit}`;
    let response: Response;
    try {
      response = await fetch(url);
    } catch (networkErr) {
      throw new Error(`Network error fetching articles: ${String(networkErr)}`);
    }
    if (!response.ok) {
      throw new Error(`Articles API returned HTTP ${response.status}`);
    }
    const data = await response.json() as { articles?: unknown };
    if (!Array.isArray(data.articles)) return [];
    return (data.articles as Array<Partial<IArticle>>).filter(
      (a): a is IArticle =>
        typeof (a as { title?: unknown })?.title === 'string' &&
        typeof (a as { url?: unknown })?.url === 'string'
    );
  }
}
