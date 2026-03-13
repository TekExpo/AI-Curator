/**
 * ============================================================================
 * AiCuratorArticleRecommender.tsx
 * ============================================================================
 *
 * Root React component for the AI Curator - Article Recommender web part.
 *
 * TAB FLOW:
 *  Tab 1 – My Interests:
 *    1. User searches for topics via the external suggest-topics API.
 *    2. Fetch current user's userPersonalization record (filter by UserId).
 *    3. Pre-select chips that match saved SelectedTags.
 *    4. On "Save My Interests" → upsert userPersonalization → switch to Tab 2.
 *
 *  Tab 2 – Recommended Articles:
 *    1. GET /_api/web/currentUser → numeric Id
 *    2. Query userPersonalization WHERE UserId eq {id}
 *    3. Extract SelectedTags → call external articles API
 *    4. Render article cards with Save & Share actions
 *
 * KEYWORDS SOURCE: Always and only from userPersonalization.SelectedTags.
 * Never from any property pane field.
 *
 * ============================================================================
 */

import * as React from 'react';
import { useState, useCallback, useEffect, useRef } from 'react';
import {
  Stack,
  Text,
  Icon,
  Link,
  Separator,
  Pivot,
  PivotItem,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  IconButton,
  TooltipHost
} from '@fluentui/react';

import { IAiCuratorArticleRecommenderProps } from './IAiCuratorArticleRecommenderProps';
import { IAiCuratorArticleRecommenderState } from './IAiCuratorArticleRecommenderState';
import { IArticle, TopicsService } from '../services/TopicsService';
import { SharePointService } from '../services/SharePointService';
import { VivaEngageService, IYammerGroup } from '../services/VivaEngageService';

import TagSelector from './TagSelector/TagSelector';
import ArticleList from './ArticleList/ArticleList';
import SharePanel from './SharePanel/SharePanel';

import styles from './AiCuratorArticleRecommender.module.scss';

// ------------------------------------------------------------------
// Constants
// ------------------------------------------------------------------

const TAB_INTERESTS = 'tab1';
const TAB_ARTICLES = 'tab2';
const TAB_SAVED = 'tab3';

const INITIAL_STATE: IAiCuratorArticleRecommenderState = {
  activeTab: TAB_ARTICLES,
  savedTags: '',
  isLoadingTags: false,
  tab1Error: '',
  tab1Success: '',
  articles: [],
  isLoadingArticles: false,
  tab2Error: '',
  tab2Info: '',
  sharePanelArticleUrl: '',
  sharePanelArticleTitle: '',
  sharePanelArticleSummary: '',
  userPersonalizationItemId: null,
  currentUserId: null,
  currentUserLoginName: '',
  savedLinks: ''
};

// ------------------------------------------------------------------
// Component
// ------------------------------------------------------------------

const AiCuratorArticleRecommender: React.FC<IAiCuratorArticleRecommenderProps> = (props) => {
  const {
    userPersonalizationListName,
    vivaEngageEnabled,
    isDarkTheme,
    hasTeamsContext,
    webPartContext
  } = props;

  // ── Services ───────────────────────────────────────────────────────────────
  const spService = useRef(new SharePointService(webPartContext));
  const topicsServiceRef = useRef(new TopicsService());

  // ── State ──────────────────────────────────────────────────────────────────
  const [state, setState] = useState<IAiCuratorArticleRecommenderState>(INITIAL_STATE);
  const [isSavingTags, setIsSavingTags] = useState(false);
  const [yammerGroups, setYammerGroups] = useState<IYammerGroup[]>([]);
  const [isLoadingGroups, setIsLoadingGroups] = useState(false);
  const [yammerGroupsError, setYammerGroupsError] = useState('');
  const [isLoadingSavedLinks, setIsLoadingSavedLinks] = useState(false);
  const [tab3Error, setTab3Error] = useState('');
  const [removingUrl, setRemovingUrl] = useState('');

  const abortControllerRef = useRef<AbortController | null>(null);

  const patchState = useCallback(
    (patch: Partial<IAiCuratorArticleRecommenderState>): void => {
      setState((prev) => ({ ...prev, ...patch }));
    },
    []
  );

  // ── Tab 1: load current user + saved tags from userPersonalization ───────────
  const loadTab1Data = useCallback(async (): Promise<void> => {
    patchState({ isLoadingTags: true, tab1Error: '', tab1Success: '' });
    try {
      const currentUser = await spService.current.getCurrentUser();
      const record = await spService.current.getUserPersonalizationByUserId(
        userPersonalizationListName,
        currentUser.Id
      );
      patchState({
        savedTags: record?.SelectedTags ?? '',
        isLoadingTags: false,
        currentUserId: currentUser.Id,
        currentUserLoginName: currentUser.LoginName,
        userPersonalizationItemId: record?.itemId ?? null,
        savedLinks: record?.SavedLinks ?? ''
      });
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      patchState({ isLoadingTags: false, tab1Error: msg });
    }
  }, [userPersonalizationListName, patchState]);

  // ── Tab 2: fetch article recommendations for current user ──────────────────
  const loadTab2Data = useCallback(async (): Promise<void> => {
    if (abortControllerRef.current) {
      abortControllerRef.current.abort();
    }
    const controller = new AbortController();
    abortControllerRef.current = controller;

    patchState({ isLoadingArticles: true, tab2Error: '', tab2Info: '', articles: [] });
    try {
      const currentUser = await spService.current.getCurrentUser();
      const record = await spService.current.getUserPersonalizationByUserId(
        userPersonalizationListName,
        currentUser.Id
      );

      if (!record) {
        if (!controller.signal.aborted) {
          patchState({
            isLoadingArticles: false,
            tab2Info:
              'No interests saved yet. Go to the My Interests tab to select your topics.'
          });
        }
        return;
      }

      const tags = record.SelectedTags?.trim() ?? '';
      if (!tags) {
        if (!controller.signal.aborted) {
          patchState({
            isLoadingArticles: false,
            tab2Info:
              'Your interests list is empty. Please select topics in the My Interests tab.'
          });
        }
        return;
      }

      patchState({
        userPersonalizationItemId: record.itemId,
        currentUserId: currentUser.Id,
        currentUserLoginName: currentUser.LoginName,
        savedLinks: record.SavedLinks ?? ''
      });

      const articles = await topicsServiceRef.current.getArticles(tags, 20);

      if (!controller.signal.aborted) {
        if (articles.length === 0) {
          patchState({
            isLoadingArticles: false,
            tab2Info: 'No articles found for your selected topics. Try updating your interests.'
          });
        } else {
          patchState({ isLoadingArticles: false, articles });
        }
      }
    } catch (err) {
      if (!controller.signal.aborted) {
        patchState({
          isLoadingArticles: false,
          tab2Error: 'Unable to fetch articles. Please try again later.'
        });
      }
    }
  }, [userPersonalizationListName, patchState]);

  // ── Tab 3: refresh SavedLinks from userPersonalization ────────────────────
  const refreshSavedLinks = useCallback(async (): Promise<void> => {
    setIsLoadingSavedLinks(true);
    setTab3Error('');
    try {
      let userId = state.currentUserId;
      if (!userId) {
        const user = await spService.current.getCurrentUser();
        userId = user.Id;
        patchState({ currentUserId: user.Id, currentUserLoginName: user.LoginName });
      }
      const record = await spService.current.getUserPersonalizationByUserId(
        userPersonalizationListName,
        userId
      );
      patchState({
        savedLinks: record?.SavedLinks ?? '',
        userPersonalizationItemId: record?.itemId ?? null
      });
    } catch (err) {
      setTab3Error(err instanceof Error ? err.message : String(err));
    } finally {
      setIsLoadingSavedLinks(false);
    }
  }, [state.currentUserId, userPersonalizationListName, patchState]);

  // ── Mount: load Tab 1 data ─────────────────────────────────────────────────
  useEffect(() => {
    loadTab2Data().catch((err) =>
      console.error('AI Curator: Failed to load Tab 2 data', err)
    );
    return () => {
      if (abortControllerRef.current) {
        abortControllerRef.current.abort();
      }
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // ── Tab switch ─────────────────────────────────────────────────────────────
  const handleTabChange = useCallback(
    (item?: PivotItem): void => {
      if (!item) return;
      const key = item.props.itemKey ?? TAB_ARTICLES;
      patchState({ activeTab: key });
      if (key === TAB_ARTICLES) {
        loadTab2Data().catch((err) =>
          console.error('AI Curator: Failed to load Tab 2 data', err)
        );
      } else if (key === TAB_SAVED) {
        refreshSavedLinks().catch((err) =>
          console.error('AI Curator: Failed to refresh saved links', err)
        );
      } else if (key === TAB_INTERESTS) {
        loadTab1Data().catch((err) =>
          console.error('AI Curator: Failed to load Tab 1 data', err)
        );
      }
    },
    [loadTab1Data, loadTab2Data, refreshSavedLinks, patchState]
  );

  // ── Save Interests (upsert userPersonalization) ────────────────────────────
  const handleSaveInterests = useCallback(async (selectedTags: string[]): Promise<void> => {
    setIsSavingTags(true);
    patchState({ tab1Error: '', tab1Success: '' });
    try {
      let userId = state.currentUserId;
      let loginName = state.currentUserLoginName;
      if (!userId) {
        const user = await spService.current.getCurrentUser();
        userId = user.Id;
        loginName = user.LoginName;
        patchState({ currentUserId: userId, currentUserLoginName: loginName });
      }
      const tagsString = selectedTags.join(', ');
      if (state.userPersonalizationItemId) {
        await spService.current.updateSelectedTags(
          userPersonalizationListName,
          state.userPersonalizationItemId,
          tagsString
        );
      } else {
        const created = await spService.current.createUserPersonalization(
          userPersonalizationListName,
          loginName,
          userId,
          tagsString
        );
        patchState({ userPersonalizationItemId: created.itemId });
      }
      patchState({ tab1Success: 'Your interests have been saved successfully!', savedTags: tagsString });
      setTimeout(() => {
        patchState({ activeTab: TAB_ARTICLES, tab1Success: '' });
        loadTab2Data().catch((err) =>
          console.error('AI Curator: Failed to load Tab 2 after save', err)
        );
      }, 1200);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      patchState({ tab1Error: `Failed to save interests: ${msg}` });
    } finally {
      setIsSavingTags(false);
    }
  }, [state.currentUserId, state.currentUserLoginName, state.userPersonalizationItemId, userPersonalizationListName, loadTab2Data, patchState]);

  // ── Save article link ──────────────────────────────────────────────────────
  const handleSaveArticle = useCallback(
    async (article: IArticle): Promise<void> => {
      const itemId = state.userPersonalizationItemId;
      if (!itemId) {
        throw new Error(
          'Could not save — user personalization record not found. ' +
            'Please save your interests first in the My Interests tab.'
        );
      }
      await spService.current.saveArticleLink(
        userPersonalizationListName,
        itemId,
        state.savedLinks,
        article.url
      );
      const updated = state.savedLinks
        ? `${state.savedLinks},${article.url}`
        : article.url;
      patchState({ savedLinks: updated });
    },
    [state.userPersonalizationItemId, state.savedLinks, userPersonalizationListName, patchState]
  );

  // ── Remove saved link ──────────────────────────────────────────────────────
  const handleRemoveSavedLink = useCallback(
    async (url: string): Promise<void> => {
      const itemId = state.userPersonalizationItemId;
      if (!itemId) return;
      setRemovingUrl(url);
      try {
        await spService.current.removeSavedLink(
          userPersonalizationListName,
          itemId,
          state.savedLinks,
          url
        );
        const updated = state.savedLinks
          .split(',')
          .map((l) => l.trim())
          .filter((l) => l.length > 0 && l !== url)
          .join(',');
        patchState({ savedLinks: updated });
      } catch (err) {
        setTab3Error(err instanceof Error ? err.message : String(err));
      } finally {
        setRemovingUrl('');
      }
    },
    [state.userPersonalizationItemId, state.savedLinks, userPersonalizationListName, patchState]
  );

  // ── Share panel ────────────────────────────────────────────────────────────
  const loadYammerGroups = useCallback((): void => {
    setIsLoadingGroups(true);
    setYammerGroupsError('');
    webPartContext.msGraphClientFactory
      .getClient('3')
      .then((graphClient) => new VivaEngageService(graphClient).getYammerGroups())
      .then((groups) => {
        setYammerGroups(groups);
        setIsLoadingGroups(false);
      })
      .catch((err) => {
        const msg = err instanceof Error ? err.message : String(err);
        setYammerGroupsError(msg);
        setIsLoadingGroups(false);
      });
  }, [webPartContext]);

  const handleShareArticle = useCallback(
    (article: IArticle): void => {
      patchState({
        sharePanelArticleUrl: article.url,
        sharePanelArticleTitle: article.title,
        sharePanelArticleSummary: article.summary ?? ''
      });
      // Always (re-)fetch groups when the panel opens so stale/failed state is cleared
      loadYammerGroups();
    },
    [loadYammerGroups, patchState]
  );

  const handlePostToVivaEngage = useCallback(
    async (groupId: string, userComments: string): Promise<void> => {
      const graphClient = await webPartContext.msGraphClientFactory.getClient('3');
      await new VivaEngageService(graphClient).postToGroup(
        groupId,
        state.sharePanelArticleUrl,
        userComments
      );
    },
    [webPartContext, state.sharePanelArticleUrl]
  );

  // ── Derived ────────────────────────────────────────────────────────────────
  const savedArticleUrls = state.savedLinks
    .split(',')
    .map((l) => l.trim())
    .filter((l) => l.length > 0);

  const sharePanelIsOpen =
    state.sharePanelArticleUrl.length > 0 && state.sharePanelArticleTitle.length > 0;

  // ── Render ─────────────────────────────────────────────────────────────────
  return (
    <section
      className={`${styles.aiCuratorArticleRecommender} ${
        hasTeamsContext ? styles.teams : ''
      } ${isDarkTheme ? styles.darkTheme : ''}`}
    >
      <Stack tokens={{ childrenGap: 16 }} className={styles.container}>
        {/* Header */}
        <Stack tokens={{ childrenGap: 4 }} className={styles.header}>
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
            <Icon iconName="LightningBolt" className={styles.headerIcon} />
            <Text variant="xLarge" className={styles.title}>
              AI Curator – Article Recommender
            </Text>
          </Stack>
        </Stack>

        <Separator />

        {/* Tabs */}
        <Pivot
          selectedKey={state.activeTab}
          onLinkClick={handleTabChange}
          styles={{
            linkIsSelected: {
              color: '#107C10',
              selectors: { '::before': { backgroundColor: '#107C10' } }
            }
          }}
        >
          <PivotItem headerText="Recommended Articles" itemKey={TAB_ARTICLES} itemIcon="Lightbulb">
            <ArticleList
              articles={state.articles}
              isLoading={state.isLoadingArticles}
              errorMessage={state.tab2Error}
              infoMessage={state.tab2Info}
              savedArticleUrls={savedArticleUrls}
              onSaveArticle={handleSaveArticle}
              onShareArticle={handleShareArticle}
              vivaEngageEnabled={vivaEngageEnabled}
            />
          </PivotItem>

          <PivotItem headerText="My Saved Links" itemKey={TAB_SAVED} itemIcon="FavoriteStar">
            {tab3Error && (
              <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline
                onDismiss={() => setTab3Error('')}
                style={{ marginTop: 8, marginBottom: 4 }}
              >
                {tab3Error}
              </MessageBar>
            )}
            {isLoadingSavedLinks ? (
              <Stack horizontalAlign="center" style={{ padding: 40 }}>
                <Spinner size={SpinnerSize.large} label="Loading saved links…" labelPosition="bottom" />
              </Stack>
            ) : savedArticleUrls.length === 0 ? (
              <Stack horizontalAlign="center" tokens={{ childrenGap: 8 }} style={{ padding: '40px 16px' }}>
                <Icon iconName="FavoriteStar" style={{ fontSize: 36, color: '#c8c6c4' }} />
                <Text variant="mediumPlus" style={{ color: '#605e5c', fontWeight: 600 }}>No saved links yet</Text>
                <Text variant="small" style={{ color: '#a19f9d', textAlign: 'center' }}>
                  Articles you save from the Recommended Articles tab will appear here.
                </Text>
              </Stack>
            ) : (
              <Stack tokens={{ childrenGap: 8 }} style={{ marginTop: 12 }}>
                <Text variant="small" style={{ color: '#605e5c', marginBottom: 4 }}>
                  {savedArticleUrls.length} saved {savedArticleUrls.length === 1 ? 'link' : 'links'}
                </Text>
                {savedArticleUrls.map((url, i) => (
                  <Stack
                    key={`saved-${i}-${url}`}
                    horizontal
                    verticalAlign="center"
                    tokens={{ childrenGap: 10 }}
                    style={{
                      padding: '10px 14px',
                      borderRadius: '8px',
                      backgroundColor: '#ffffff',
                      border: '1px solid #edebe9',
                      boxShadow: '0 1px 3px rgba(0,0,0,0.08)'
                    }}
                  >
                    <Icon iconName="Link" style={{ fontSize: 16, color: '#107C10', flexShrink: 0 }} />
                    <Link
                      href={url}
                      target="_blank"
                      rel="noopener noreferrer"
                      style={{ flexGrow: 1, wordBreak: 'break-all', color: '#107C10', fontSize: 13 }}
                    >
                      {url}
                    </Link>
                    <TooltipHost content="Remove link">
                      <IconButton
                        iconProps={{ iconName: 'Delete' }}
                        ariaLabel="Remove saved link"
                        disabled={removingUrl === url}
                        onClick={() => { void handleRemoveSavedLink(url); }}
                        styles={{
                          root: { color: '#605e5c', flexShrink: 0 },
                          rootHovered: { color: '#a4262c' },
                          icon: { fontSize: 14 }
                        }}
                      />
                    </TooltipHost>
                  </Stack>
                ))}
              </Stack>
            )}
          </PivotItem>

          <PivotItem headerText="My Interests" itemKey={TAB_INTERESTS} itemIcon="Tag">
            {state.tab1Error && (
              <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline
                onDismiss={() => patchState({ tab1Error: '' })}
                style={{ marginTop: 8, marginBottom: 4 }}
              >
                {state.tab1Error}
              </MessageBar>
            )}
            {state.isLoadingTags ? (
              <Stack horizontalAlign="center" style={{ padding: 40 }}>
                <Spinner size={SpinnerSize.large} label="Loading your interests…" labelPosition="bottom" />
              </Stack>
            ) : (
              <TagSelector
                savedTags={state.savedTags}
                successMessage={state.tab1Success}
                onSave={(tags) => { void handleSaveInterests(tags); }}
                isSaving={isSavingTags}
              />
            )}
          </PivotItem>
        </Pivot>
      </Stack>

      <SharePanel
        isOpen={sharePanelIsOpen}
        articleUrl={state.sharePanelArticleUrl}
        articleTitle={state.sharePanelArticleTitle}
        articleSummary={state.sharePanelArticleSummary}
        groups={yammerGroups}
        isLoadingGroups={isLoadingGroups}
        groupLoadError={yammerGroupsError}
        onRetryLoadGroups={loadYammerGroups}
      onDismiss={() => patchState({ sharePanelArticleUrl: '', sharePanelArticleTitle: '', sharePanelArticleSummary: '' })}
        onPost={handlePostToVivaEngage}
      />
    </section>
  );
};

export default AiCuratorArticleRecommender;
