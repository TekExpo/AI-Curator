/**
 * ============================================================================
 * AiCuratorArticleRecommender.tsx
 * ============================================================================
 *
 * Main React functional component for the AI Curator - Article Recommender
 * web part. Uses React hooks (useState, useEffect, useCallback, useRef) for
 * state management and lifecycle.
 *
 * FLOW:
 *  1. On mount (and when relevant props change), fetch items from the
 *     configured SharePoint list using @pnp/sp.
 *  2. Extract and concatenate keywords from the configured column.
 *  3. POST the keywords + maxResults to the LLM endpoint using native fetch.
 *  4. Parse the JSON response and render article cards with Fluent UI.
 *
 * CACHING:
 *  When enableCaching is true, LLM responses are stored in sessionStorage
 *  keyed by a hash of the keywords + maxResults. This avoids redundant API
 *  calls when revisiting the same page or when the web part re-renders
 *  without property changes.
 *
 * EXTENDING THE LLM PAYLOAD:
 *  To add additional fields to the LLM request body:
 *    1. Add a new property in IAiCuratorArticleRecommenderWebPartProps
 *    2. Add a corresponding PropertyPane field in getPropertyPaneConfiguration()
 *    3. Pass it through to this component via IAiCuratorArticleRecommenderProps
 *    4. Include it in the `body` object inside callLlmEndpoint()
 *
 * ============================================================================
 */

import * as React from 'react';
import { useState, useEffect, useCallback, useRef } from 'react';
import {
  Stack,
  Text,
  Link,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Icon,
  Separator,
  IStackTokens,
  IStackItemStyles,
  mergeStyles
} from '@fluentui/react';

import {
  IAiCuratorArticleRecommenderProps,
  IArticleRecommendation
} from './IAiCuratorArticleRecommenderProps';

import styles from './AiCuratorArticleRecommender.module.scss';

// ------------------------------------------------------------------
// Constants
// ------------------------------------------------------------------

const CACHE_KEY_PREFIX = 'ai-curator-cache-';
const CACHE_TTL_MS = 15 * 60 * 1000; // 15 minutes

const stackTokens: IStackTokens = { childrenGap: 12 };
const articleCardStyles: IStackItemStyles = {
  root: {
    padding: '12px 16px',
    borderRadius: '8px',
    boxShadow: '0 1px 3px rgba(0,0,0,0.12), 0 1px 2px rgba(0,0,0,0.06)',
    transition: 'box-shadow 0.2s ease, transform 0.15s ease',
    backgroundColor: '#ffffff',
    selectors: {
      ':hover': {
        boxShadow: '0 4px 12px rgba(0,0,0,0.15), 0 2px 4px rgba(0,0,0,0.10)',
        transform: 'translateY(-1px)'
      }
    }
  }
};

// ------------------------------------------------------------------
// Utility Functions
// ------------------------------------------------------------------

/**
 * Generates a simple hash string for cache key derivation.
 * Uses a basic DJB2-style hash – not cryptographic, just for keying.
 */
function hashString(input: string): string {
  let hash = 5381;
  for (let i = 0; i < input.length; i++) {
    // eslint-disable-next-line no-bitwise
    hash = ((hash << 5) + hash) + input.charCodeAt(i);
  }
  return hash.toString(36);
}

/**
 * Validates that a parsed JSON value conforms to the expected LLM response
 * schema: an array of objects each with title, url, and summary strings.
 */
function isValidArticleArray(data: unknown): data is IArticleRecommendation[] {
  if (!Array.isArray(data)) return false;
  return data.every(
    (item: Record<string, unknown>) =>
      typeof item === 'object' &&
      item !== null &&
      typeof item.title === 'string' &&
      typeof item.url === 'string' &&
      typeof item.summary === 'string'
  );
}

// ------------------------------------------------------------------
// Component
// ------------------------------------------------------------------

const AiCuratorArticleRecommender: React.FC<IAiCuratorArticleRecommenderProps> = (props) => {
  const {
    llmEndpointUrl,
    listName,
    keywordColumnName,
    siteUrl,
    maxArticles,
    enableCaching,
    spInstance,
    isDarkTheme,
    hasTeamsContext
  } = props;

  // ----- State -----
  const [articles, setArticles] = useState<IArticleRecommendation[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [errorMessage, setErrorMessage] = useState<string>('');
  const [keywords, setKeywords] = useState<string>('');

  // Abort controller ref for cancelling in-flight requests on unmount/re-fetch
  const abortControllerRef = useRef<AbortController | null>(null);

  // ----- Fetch keywords from SharePoint -----
  const fetchKeywordsFromList = useCallback(async (): Promise<string> => {
    try {
      // Determine which SP site to query.
      // If siteUrl is provided, use it; otherwise, use the current context site.
      let spQuery = spInstance;

      if (siteUrl && siteUrl.trim().length > 0) {
        // NOTE: When targetting a different site, PnPjs can use .using(SPFx(...))
        // with the web URL. For simplicity here we show a comment – you would
        // need to create a new spfi instance pointing to the other site.
        // For cross-site queries, consider: spfi(siteUrl).using(SPFx(context))
        // For now we use the current context SP instance.
        console.warn(
          'AI Curator: Custom siteUrl is set but cross-site PnPjs setup is not configured. ' +
          'Using current site context. See code comments for cross-site setup.'
        );
      }

      // Fetch all items from the list, selecting only the keyword column.
      // We retrieve up to 5000 items (SharePoint list view threshold).
      const items: Record<string, unknown>[] = await spQuery.web.lists
        .getByTitle(listName)
        .items
        .select(keywordColumnName)
        .top(5000)();

      if (!items || items.length === 0) {
        return '';
      }

      /**
       * KEYWORD AGGREGATION STRATEGY:
       *
       * We concatenate keywords from ALL items in the list, separated by commas.
       * This provides the broadest context to the LLM for generating recommendations.
       *
       * Alternative approaches:
       *   - Use only the FIRST item's keywords (for single-source scenarios)
       *   - Use only the MOST RECENT item (sort by Modified desc, top 1)
       *   - Deduplicate keywords before sending
       *
       * The current approach (all items) is best for lists where each item
       * represents a topic or category the user is interested in.
       */
      const allKeywords = items
        .map((item) => {
          const value = item[keywordColumnName];
          return typeof value === 'string' ? value.trim() : '';
        })
        .filter((kw) => kw.length > 0)
        .join(', ');

      // Deduplicate keywords to reduce payload size
      const uniqueKeywords = [...new Set(
        allKeywords
          .split(',')
          .map((k) => k.trim().toLowerCase())
          .filter((k) => k.length > 0)
      )];

      return uniqueKeywords.join(', ');
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      throw new Error(`Failed to fetch keywords from list "${listName}": ${message}`);
    }
  }, [spInstance, listName, keywordColumnName, siteUrl]);

  // ----- Call LLM Endpoint -----
  const callLlmEndpoint = useCallback(async (
    kwString: string,
    signal: AbortSignal
  ): Promise<IArticleRecommendation[]> => {
    if (!llmEndpointUrl || llmEndpointUrl.trim().length === 0) {
      throw new Error('LLM Endpoint URL is not configured. Please set it in the web part properties.');
    }

    // Check cache first
    if (enableCaching) {
      const cacheKey = CACHE_KEY_PREFIX + hashString(`${kwString}|${maxArticles}`);
      try {
        const cached = sessionStorage.getItem(cacheKey);
        if (cached) {
          const parsed = JSON.parse(cached) as { timestamp: number; data: IArticleRecommendation[] };
          if (Date.now() - parsed.timestamp < CACHE_TTL_MS) {
            return parsed.data;
          }
          // Cache expired – remove stale entry
          sessionStorage.removeItem(cacheKey);
        }
      } catch {
        // Cache read failed – proceed with fresh fetch
      }
    }

    /**
     * LLM REQUEST PAYLOAD:
     *
     * The expected format is:
     *   { "keywords": "keyword1, keyword2, ...", "maxResults": 5 }
     *
     * TO EXTEND: Add additional fields here if your LLM endpoint requires
     * them (e.g., "language", "userId", "category"). Remember to also add
     * corresponding Property Pane fields and props.
     */
    const payload = {
      keywords: kwString,
      maxResults: maxArticles
    };

    const response = await fetch(llmEndpointUrl.trim(), {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Accept': 'application/json'
        // ---------------------------------------------------------------
        // AUTHENTICATION:
        // If your LLM endpoint requires a Bearer token, add it here:
        //   'Authorization': `Bearer ${yourTokenVariable}`
        //
        // For Azure AD-protected endpoints, you can acquire a token using:
        //   this.context.aadTokenProviderFactory
        //     .getTokenProvider()
        //     .then(provider => provider.getToken('your-resource-id'))
        //
        // Pass the token through props from the web part class.
        // ---------------------------------------------------------------
      },
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

    if (!isValidArticleArray(data)) {
      throw new Error(
        'LLM response does not match expected schema. Expected: ' +
        '[{ "title": string, "url": string, "summary": string }, ...]. ' +
        `Received: ${JSON.stringify(data).substring(0, 300)}`
      );
    }

    // Store in cache
    if (enableCaching) {
      const cacheKey = CACHE_KEY_PREFIX + hashString(`${kwString}|${maxArticles}`);
      try {
        sessionStorage.setItem(cacheKey, JSON.stringify({
          timestamp: Date.now(),
          data
        }));
      } catch {
        // sessionStorage full or unavailable – silently continue
      }
    }

    return data;
  }, [llmEndpointUrl, maxArticles, enableCaching]);

  // ----- Main Data Pipeline -----
  const loadRecommendations = useCallback(async (): Promise<void> => {
    // Cancel any previous in-flight request
    if (abortControllerRef.current) {
      abortControllerRef.current.abort();
    }

    const controller = new AbortController();
    abortControllerRef.current = controller;

    setIsLoading(true);
    setErrorMessage('');
    setArticles([]);

    try {
      // Step 1: Fetch keywords from SharePoint list
      const kw = await fetchKeywordsFromList();
      setKeywords(kw);

      if (!kw || kw.trim().length === 0) {
        setErrorMessage(
          `No keywords found in the "${listName}" list (column: "${keywordColumnName}"). ` +
          'Please add items with keywords to get recommendations.'
        );
        setIsLoading(false);
        return;
      }

      // Step 2: Call LLM endpoint with extracted keywords
      const recommendedArticles = await callLlmEndpoint(kw, controller.signal);

      // Only update state if this request wasn't aborted
      if (!controller.signal.aborted) {
        setArticles(recommendedArticles);
        setIsLoading(false);
      }
    } catch (err) {
      if (!controller.signal.aborted) {
        const msg = err instanceof Error ? err.message : String(err);
        setErrorMessage(msg);
        setIsLoading(false);
      }
    }
  }, [fetchKeywordsFromList, callLlmEndpoint, listName, keywordColumnName]);

  // ----- Effect: trigger load on mount and when relevant properties change -----
  useEffect(() => {
    // Only attempt to load if we have the minimum required configuration
    if (listName && keywordColumnName) {
      loadRecommendations().catch((err) => {
        console.error('AI Curator: Failed to load recommendations', err);
      });
    }

    // Cleanup: abort in-flight requests when component unmounts or deps change
    return () => {
      if (abortControllerRef.current) {
        abortControllerRef.current.abort();
      }
    };
  }, [llmEndpointUrl, listName, keywordColumnName, siteUrl, maxArticles, enableCaching]);

  // ----- Render Helpers -----

  /** Renders the header section with title and optional keyword summary */
  const renderHeader = (): React.ReactElement => (
    <Stack tokens={{ childrenGap: 4 }} className={styles.header}>
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
        <Icon
          iconName="LightningBolt"
          className={styles.headerIcon}
        />
        <Text variant="xLarge" className={styles.title}>
          AI Curator – Article Recommender
        </Text>
      </Stack>
      {keywords && (
        <Text variant="small" className={styles.keywordsLabel}>
          Keywords: <em>{keywords.length > 120 ? keywords.substring(0, 120) + '…' : keywords}</em>
        </Text>
      )}
    </Stack>
  );

  /** Renders the loading spinner */
  const renderLoading = (): React.ReactElement => (
    <Stack horizontalAlign="center" verticalAlign="center" className={styles.loadingContainer}>
      <Spinner
        size={SpinnerSize.large}
        label="Fetching article recommendations…"
        labelPosition="bottom"
      />
    </Stack>
  );

  /** Renders an error message bar */
  const renderError = (): React.ReactElement => (
    <MessageBar
      messageBarType={MessageBarType.error}
      isMultiline
      dismissButtonAriaLabel="Close"
      onDismiss={() => setErrorMessage('')}
      className={styles.errorBar}
    >
      <Text variant="smallPlus">{errorMessage}</Text>
    </MessageBar>
  );

  /** Renders the "no results" state */
  const renderEmpty = (): React.ReactElement => (
    <MessageBar
      messageBarType={MessageBarType.info}
      isMultiline={false}
      className={styles.infoBar}
    >
      No article recommendations were returned. Try adjusting the keywords in your SharePoint list.
    </MessageBar>
  );

  /** Renders the configuration prompt when LLM URL is not set */
  const renderConfigPrompt = (): React.ReactElement => (
    <Stack horizontalAlign="center" verticalAlign="center" className={styles.configPrompt}>
      <Icon iconName="Settings" className={styles.configIcon} />
      <Text variant="large">Configure the web part</Text>
      <Text variant="medium" className={styles.configText}>
        Open the property pane (edit icon) and set the LLM Endpoint URL,
        SharePoint list name, and keyword column to get started.
      </Text>
    </Stack>
  );

  /** Renders a single article recommendation card */
  const renderArticleCard = (
    article: IArticleRecommendation,
    index: number
  ): React.ReactElement => (
    <Stack
      key={`article-${index}-${article.url}`}
      styles={articleCardStyles}
      className={styles.articleCard}
    >
      <Stack horizontal verticalAlign="start" tokens={{ childrenGap: 10 }}>
        <Icon
          iconName="TextDocument"
          className={styles.articleIcon}
        />
        <Stack tokens={{ childrenGap: 4 }} grow>
          <Link
            href={article.url}
            target="_blank"
            rel="noopener noreferrer"
            className={styles.articleLink}
            title={`Open "${article.title}" in a new tab`}
          >
            {article.title}
          </Link>
          {article.summary && (
            <Text variant="small" className={styles.articleSummary}>
              {article.summary}
            </Text>
          )}
          <Text variant="tiny" className={styles.articleUrl}>
            {article.url.length > 80
              ? article.url.substring(0, 77) + '…'
              : article.url}
          </Text>
        </Stack>
      </Stack>
    </Stack>
  );

  /** Renders the list of article recommendation cards */
  const renderArticles = (): React.ReactElement => (
    <Stack tokens={stackTokens} className={styles.articleList}>
      {articles.map((article, index) => renderArticleCard(article, index))}
      <Text variant="tiny" className={styles.footerNote}>
        Showing {articles.length} recommendation{articles.length !== 1 ? 's' : ''} • Powered by AI
      </Text>
    </Stack>
  );

  // ----- Main Render -----
  return (
    <section
      className={`${styles.aiCuratorArticleRecommender} ${
        hasTeamsContext ? styles.teams : ''
      } ${isDarkTheme ? styles.darkTheme : ''}`}
    >
      <Stack tokens={{ childrenGap: 16 }} className={styles.container}>
        {renderHeader()}
        <Separator />

        {/* Show config prompt if no LLM URL is set */}
        {!llmEndpointUrl && !isLoading && !errorMessage && renderConfigPrompt()}

        {/* Loading state */}
        {isLoading && renderLoading()}

        {/* Error state */}
        {errorMessage && !isLoading && renderError()}

        {/* Empty results state */}
        {!isLoading && !errorMessage && llmEndpointUrl && articles.length === 0 && keywords && renderEmpty()}

        {/* Article results */}
        {!isLoading && !errorMessage && articles.length > 0 && renderArticles()}
      </Stack>
    </section>
  );
};

export default AiCuratorArticleRecommender;
