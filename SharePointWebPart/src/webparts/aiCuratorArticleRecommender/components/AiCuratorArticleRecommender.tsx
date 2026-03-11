/**
 * ============================================================================
 * AiCuratorArticleRecommender.tsx
 * ============================================================================
 *
 * Root React component for the AI Curator - Article Recommender web part.
 *
 * TAB FLOW:
 *  Tab 1 – My Interests:
 *    1. Fetch all tags from Articles list (Keywords column)
 *    2. Fetch current user's userPersonalization record (filter by UserId)
 *    3. Pre-select tags stored in SelectedTags
 *    4. On "Save My Interests" → upsert userPersonalization → switch to Tab 2
 *
 *  Tab 2 – Recommended Articles:
 *    1. GET /_api/web/currentUser → numeric Id
 *    2. Query userPersonalization WHERE UserId eq {id}
 *    3. Extract SelectedTags → pass directly to OpenAI service
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
  Separator,
  Pivot,
  PivotItem,
  MessageBar,
  MessageBarType
} from '@fluentui/react';

import { IAiCuratorArticleRecommenderProps } from './IAiCuratorArticleRecommenderProps';
import { IAiCuratorArticleRecommenderState } from './IAiCuratorArticleRecommenderState';
import { IArticle, OpenAIService } from '../services/OpenAIService';
import { SharePointService } from '../services/SharePointService';
import { VivaEngageService, IYammerGroup } from '../services/VivaEngageService';
import { aiCuratorArticleRecommenderConfig } from '../AiCuratorArticleRecommender.config';

import TagSelector from './TagSelector/TagSelector';
import ArticleList from './ArticleList/ArticleList';
import SharePanel from './SharePanel/SharePanel';

import styles from './AiCuratorArticleRecommender.module.scss';

// ------------------------------------------------------------------
// Constants
// ------------------------------------------------------------------

const TAB_INTERESTS = 'tab1';
const TAB_ARTICLES = 'tab2';

const INITIAL_STATE: IAiCuratorArticleRecommenderState = {
  activeTab: TAB_INTERESTS,
  availableTags: [],
  selectedTags: [],
  isLoadingTags: false,
  tab1Error: '',
  tab1Success: '',
  articles: [],
  isLoadingArticles: false,
  tab2Error: '',
  tab2Info: '',
  sharePanelArticleUrl: '',
  sharePanelArticleTitle: '',
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
    openAiModel,
    openAiSystemPrompt,
    articlesListName,
    maxArticles,
    enableCaching,
    userPersonalizationListName,
    vivaEngageEnabled,
    isDarkTheme,
    hasTeamsContext,
    webPartContext
  } = props;

  // ── Services ───────────────────────────────────────────────────────────────
  const spService = useRef(new SharePointService(webPartContext));
  const openAiServiceRef = useRef(
    new OpenAIService(
      aiCuratorArticleRecommenderConfig.llmEndpointUrl,
      aiCuratorArticleRecommenderConfig.openAiApiKey
    )
  );

  // ── State ──────────────────────────────────────────────────────────────────
  const [state, setState] = useState<IAiCuratorArticleRecommenderState>(INITIAL_STATE);
  const [isSavingTags, setIsSavingTags] = useState(false);
  const [yammerGroups, setYammerGroups] = useState<IYammerGroup[]>([]);
  const [isLoadingGroups, setIsLoadingGroups] = useState(false);

  const abortControllerRef = useRef<AbortController | null>(null);

  const patchState = useCallback(
    (patch: Partial<IAiCuratorArticleRecommenderState>): void => {
      setState((prev) => ({ ...prev, ...patch }));
    },
    []
  );

  // ── Tab 1: load available tags + pre-select existing user tags ──────────────
  const loadTab1Data = useCallback(async (): Promise<void> => {
    patchState({ isLoadingTags: true, tab1Error: '', tab1Success: '' });
    try {
      const [availableTags, currentUser] = await Promise.all([
        spService.current.getTagsFromArticlesList(articlesListName),
        spService.current.getCurrentUser()
      ]);
      const record = await spService.current.getUserPersonalizationByUserId(
        userPersonalizationListName,
        currentUser.Id
      );
      const preSelected =
        record?.SelectedTags
          ?.split(',')
          .map((t) => t.trim().toLowerCase())
          .filter((t) => t.length > 0) ?? [];
      patchState({
        availableTags,
        selectedTags: preSelected,
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
  }, [articlesListName, userPersonalizationListName, patchState]);

  // ── Tab 2: fetch OpenAI recommendations for current user ───────────────────
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
              'No interests saved yet. Go to the My Interests tab to select your tags.'
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
              'Your interests list is empty. Please select tags in the My Interests tab.'
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

      const articles = await openAiServiceRef.current.getArticleRecommendations(
        tags,
        openAiModel,
        openAiSystemPrompt,
        maxArticles,
        enableCaching,
        controller.signal
      );

      if (!controller.signal.aborted) {
        patchState({ isLoadingArticles: false, articles });
      }
    } catch (err) {
      if (!controller.signal.aborted) {
        const msg = err instanceof Error ? err.message : String(err);
        patchState({ isLoadingArticles: false, tab2Error: msg });
      }
    }
  }, [userPersonalizationListName, openAiModel, openAiSystemPrompt, maxArticles, enableCaching, patchState]);

  // ── Mount: load Tab 1 data ─────────────────────────────────────────────────
  useEffect(() => {
    loadTab1Data().catch((err) =>
      console.error('AI Curator: Failed to load Tab 1 data', err)
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
      const key = item.props.itemKey ?? TAB_INTERESTS;
      patchState({ activeTab: key });
      if (key === TAB_ARTICLES) {
        loadTab2Data().catch((err) =>
          console.error('AI Curator: Failed to load Tab 2 data', err)
        );
      }
    },
    [loadTab2Data, patchState]
  );

  // ── Tag chip toggle ────────────────────────────────────────────────────────
  const handleTagToggle = useCallback((tag: string): void => {
    setState((prev) => {
      const isSelected = prev.selectedTags.includes(tag);
      return {
        ...prev,
        selectedTags: isSelected
          ? prev.selectedTags.filter((t) => t !== tag)
          : [...prev.selectedTags, tag]
      };
    });
  }, []);

  // ── Save Interests (upsert userPersonalization) ────────────────────────────
  const handleSaveInterests = useCallback(async (): Promise<void> => {
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
      const tagsString = state.selectedTags.join(', ');
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
      patchState({ tab1Success: 'Your interests have been saved successfully!' });
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
  }, [state, userPersonalizationListName, loadTab2Data, patchState]);

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

  // ── Share panel ────────────────────────────────────────────────────────────
  const handleShareArticle = useCallback(
    (article: IArticle): void => {
      patchState({
        sharePanelArticleUrl: article.url,
        sharePanelArticleTitle: article.title
      });
      if (vivaEngageEnabled && yammerGroups.length === 0 && !isLoadingGroups) {
        setIsLoadingGroups(true);
        webPartContext.msGraphClientFactory
          .getClient('3')
          .then((graphClient) => new VivaEngageService(graphClient).getYammerGroups())
          .then((groups) => {
            setYammerGroups(groups);
            setIsLoadingGroups(false);
          })
          .catch((err) => {
            console.error('AI Curator: Failed to fetch Viva Engage groups', err);
            setIsLoadingGroups(false);
          });
      }
    },
    [vivaEngageEnabled, yammerGroups.length, isLoadingGroups, webPartContext, patchState]
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
            <TagSelector
              availableTags={state.availableTags}
              selectedTags={state.selectedTags}
              isLoading={state.isLoadingTags}
              errorMessage={''}
              successMessage={state.tab1Success}
              onTagToggle={handleTagToggle}
              onSave={() => { void handleSaveInterests(); }}
              isSaving={isSavingTags}
            />
          </PivotItem>

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
        </Pivot>
      </Stack>

      <SharePanel
        isOpen={sharePanelIsOpen}
        articleUrl={state.sharePanelArticleUrl}
        articleTitle={state.sharePanelArticleTitle}
        groups={yammerGroups}
        isLoadingGroups={isLoadingGroups}
        onDismiss={() => patchState({ sharePanelArticleUrl: '', sharePanelArticleTitle: '' })}
        onPost={handlePostToVivaEngage}
      />
    </section>
  );
};

export default AiCuratorArticleRecommender;
