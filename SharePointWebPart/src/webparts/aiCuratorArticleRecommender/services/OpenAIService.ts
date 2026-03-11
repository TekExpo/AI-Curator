export interface IArticle {
  title: string;
  description: string;
  url: string;
}

// ------------------------------------------------------------------
// Internal helpers (moved from AiCuratorArticleRecommender.tsx)
// ------------------------------------------------------------------

const CACHE_KEY_PREFIX = 'ai-curator-cache-';
const CACHE_TTL_MS = 15 * 60 * 1000; // 15 minutes

const DEFAULT_OPENAI_SYSTEM_PROMPT = [
  'You recommend articles based on a list of keywords.',
  'Return only valid JSON.',
  'The JSON must be an array of objects with this exact schema:',
  '[{"title":"string","url":"https://example.com","summary":"string"}]',
  'Use real-looking article titles and absolute URLs.',
  'Return at most {{maxArticles}} items.'
].join(' ');

function hashString(input: string): string {
  let hash = 5381;
  for (let i = 0; i < input.length; i++) {
    // eslint-disable-next-line no-bitwise
    hash = ((hash << 5) + hash) + input.charCodeAt(i);
  }
  return hash.toString(36);
}

function isValidArticleArray(data: unknown): data is IArticle[] {
  if (!Array.isArray(data)) return false;
  return data.every(
    (item) =>
      typeof item === 'object' &&
      item !== null &&
      typeof (item as Record<string, unknown>).title === 'string' &&
      typeof (item as Record<string, unknown>).url === 'string' &&
      (typeof (item as Record<string, unknown>).summary === 'string' ||
        typeof (item as Record<string, unknown>).description === 'string')
  );
}

function extractJsonPayload(text: string): string {
  const fencedMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const candidate = fencedMatch?.[1]?.trim() ?? text.trim();
  const arrayStart = candidate.indexOf('[');
  const arrayEnd = candidate.lastIndexOf(']');
  if (arrayStart >= 0 && arrayEnd > arrayStart) {
    return candidate.substring(arrayStart, arrayEnd + 1);
  }
  return candidate;
}

function normalizeArticleArray(raw: unknown[]): IArticle[] {
  return raw.map((item) => {
    const r = item as Record<string, unknown>;
    return {
      title: String(r.title ?? ''),
      url: String(r.url ?? ''),
      description: String(r.summary ?? r.description ?? '')
    };
  });
}

function parseArticleArrayFromText(text: string): IArticle[] | null {
  try {
    const parsed = JSON.parse(extractJsonPayload(text)) as unknown;
    if (isValidArticleArray(parsed)) return normalizeArticleArray(parsed);
    if (
      typeof parsed === 'object' && parsed !== null && 'articles' in parsed
    ) {
      const articles = (parsed as { articles?: unknown }).articles;
      if (isValidArticleArray(articles)) return normalizeArticleArray(articles);
    }
  } catch {
    // ignore
  }
  return null;
}

function parseOpenAiResponse(data: unknown): IArticle[] | null {
  if (isValidArticleArray(data)) return normalizeArticleArray(data);

  if (typeof data === 'object' && data !== null && 'articles' in data) {
    const articles = (data as { articles?: unknown }).articles;
    if (isValidArticleArray(articles)) return normalizeArticleArray(articles);
  }

  const chatContent = (data as {
    choices?: Array<{ message?: { content?: unknown } }>;
  } | null)?.choices?.[0]?.message?.content;

  if (typeof chatContent === 'string') return parseArticleArrayFromText(chatContent);

  if (Array.isArray(chatContent)) {
    const contentText = chatContent
      .map((part) => {
        if (typeof part === 'string') return part;
        if (typeof part === 'object' && part !== null && 'text' in part) {
          const t = (part as { text?: unknown }).text;
          return typeof t === 'string' ? t : '';
        }
        return '';
      })
      .join('');
    if (contentText) return parseArticleArrayFromText(contentText);
  }

  const responsesText = (data as {
    output?: Array<{ content?: Array<{ text?: string }> }>;
  } | null)?.output?.[0]?.content?.map((part) => part.text ?? '').join('');
  if (responsesText) return parseArticleArrayFromText(responsesText);

  return null;
}

// ------------------------------------------------------------------
// OpenAIService
// ------------------------------------------------------------------

/**
 * Handles all OpenAI API communication.
 * Keywords are always the raw SelectedTags string from userPersonalization.
 */
export class OpenAIService {
  private readonly _endpointUrl: string;
  private readonly _apiKey: string;

  constructor(endpointUrl: string, apiKey: string) {
    this._endpointUrl = endpointUrl;
    this._apiKey = apiKey;
  }

  public async getArticleRecommendations(
    tags: string,
    model: string,
    systemPrompt: string,
    maxArticles: number,
    enableCaching: boolean,
    signal?: AbortSignal
  ): Promise<IArticle[]> {
    const endpointUrl = this._endpointUrl?.trim();
    if (!endpointUrl) {
      throw new Error(
        'LLM Endpoint URL is not configured. Please set it in the web part config file.'
      );
    }

    const resolvedPrompt = (systemPrompt?.trim().length > 0 ? systemPrompt.trim() : DEFAULT_OPENAI_SYSTEM_PROMPT)
      .replace(/\{\{keywords\}\}/g, tags)
      .replace(/\{\{maxArticles\}\}/g, String(maxArticles));

    if (enableCaching) {
      const cacheKey =
        CACHE_KEY_PREFIX +
        hashString(`${endpointUrl}|${model}|${resolvedPrompt}|${tags}|${maxArticles}`);
      try {
        const cached = sessionStorage.getItem(cacheKey);
        if (cached) {
          const parsed = JSON.parse(cached) as { timestamp: number; data: IArticle[] };
          if (Date.now() - parsed.timestamp < CACHE_TTL_MS) {
            return parsed.data;
          }
          sessionStorage.removeItem(cacheKey);
        }
      } catch {
        // sessionStorage unavailable — continue
      }
    }

    const isOpenAiEndpoint =
      /api\.openai\.com|openai\.azure\.com/i.test(endpointUrl) ||
      endpointUrl.includes('/chat/completions') ||
      endpointUrl.includes('/responses');

    const userPrompt = [
      `Keywords: ${tags}`,
      `Maximum articles: ${maxArticles}`,
      'Return article recommendations for these keywords with working URLs and concise summaries.'
    ].join('\n');

    const payload: Record<string, unknown> = isOpenAiEndpoint
      ? {
          messages: [
            { role: 'system', content: resolvedPrompt },
            { role: 'user', content: userPrompt }
          ],
          temperature: 0.2,
          ...(model?.trim().length > 0 ? { model: model.trim() } : {})
        }
      : { keywords: tags, maxResults: maxArticles, prompt: resolvedPrompt };

    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      Accept: 'application/json'
    };
    if (this._apiKey?.trim().length > 0) {
      if (/openai\.azure\.com/i.test(endpointUrl)) {
        headers['api-key'] = this._apiKey.trim();
      } else {
        headers.Authorization = `Bearer ${this._apiKey.trim()}`;
      }
    }

    const response = await fetch(endpointUrl, {
      method: 'POST',
      headers,
      body: JSON.stringify(payload),
      signal
    });

    if (!response.ok) {
      const errorBody = await response.text().catch(() => 'No response body');
      throw new Error(
        `LLM endpoint returned HTTP ${response.status}: ${response.statusText}. ` +
          `Body: ${errorBody.substring(0, 200)}`
      );
    }

    let data: unknown;
    try {
      data = await response.json();
    } catch {
      throw new Error(
        'LLM endpoint returned invalid JSON. Expected an array of ' +
          '{ title, url, summary } objects.'
      );
    }

    const parsedArticles = parseOpenAiResponse(data);
    if (!parsedArticles) {
      throw new Error(
        'LLM response does not match expected schema. Expected array of ' +
          '[{ "title": string, "url": string, "summary": string }, ...]. ' +
          `Received: ${JSON.stringify(data).substring(0, 300)}`
      );
    }

    if (enableCaching) {
      const cacheKey =
        CACHE_KEY_PREFIX +
        hashString(`${endpointUrl}|${model}|${resolvedPrompt}|${tags}|${maxArticles}`);
      try {
        sessionStorage.setItem(
          cacheKey,
          JSON.stringify({ timestamp: Date.now(), data: parsedArticles })
        );
      } catch {
        // sessionStorage full or unavailable
      }
    }

    return parsedArticles;
  }
}
