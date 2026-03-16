"use strict";
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
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var React = tslib_1.__importStar(require("react"));
var react_1 = require("react");
var react_2 = require("@fluentui/react");
var TopicsService_1 = require("../services/TopicsService");
var SharePointService_1 = require("../services/SharePointService");
var VivaEngageService_1 = require("../services/VivaEngageService");
var TagSelector_1 = tslib_1.__importDefault(require("./TagSelector/TagSelector"));
var ArticleList_1 = tslib_1.__importDefault(require("./ArticleList/ArticleList"));
var SharePanel_1 = tslib_1.__importDefault(require("./SharePanel/SharePanel"));
var LinkedInPanel_1 = tslib_1.__importDefault(require("./LinkedInPanel/LinkedInPanel"));
var AiCuratorArticleRecommender_module_scss_1 = tslib_1.__importDefault(require("./AiCuratorArticleRecommender.module.scss"));
// ------------------------------------------------------------------
// Constants
// ------------------------------------------------------------------
var TAB_INTERESTS = 'tab1';
var TAB_ARTICLES = 'tab2';
var TAB_SAVED = 'tab3';
var INITIAL_STATE = {
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
    linkedInPanelArticleUrl: '',
    linkedInPanelArticleTitle: '',
    linkedInPanelArticleSummary: '',
    userPersonalizationItemId: null,
    currentUserId: null,
    currentUserLoginName: '',
    savedLinks: ''
};
// ------------------------------------------------------------------
// Component
// ------------------------------------------------------------------
var AiCuratorArticleRecommender = function (props) {
    var userPersonalizationListName = props.userPersonalizationListName, vivaEngageEnabled = props.vivaEngageEnabled, isDarkTheme = props.isDarkTheme, hasTeamsContext = props.hasTeamsContext, webPartContext = props.webPartContext;
    // ── Services ───────────────────────────────────────────────────────────────
    var spService = (0, react_1.useRef)(new SharePointService_1.SharePointService(webPartContext));
    var topicsServiceRef = (0, react_1.useRef)(new TopicsService_1.TopicsService());
    // ── State ──────────────────────────────────────────────────────────────────
    var _a = (0, react_1.useState)(INITIAL_STATE), state = _a[0], setState = _a[1];
    var _b = (0, react_1.useState)(false), isSavingTags = _b[0], setIsSavingTags = _b[1];
    var _c = (0, react_1.useState)([]), yammerGroups = _c[0], setYammerGroups = _c[1];
    var _d = (0, react_1.useState)(false), isLoadingGroups = _d[0], setIsLoadingGroups = _d[1];
    var _e = (0, react_1.useState)(''), yammerGroupsError = _e[0], setYammerGroupsError = _e[1];
    var _f = (0, react_1.useState)(false), isLoadingSavedLinks = _f[0], setIsLoadingSavedLinks = _f[1];
    var _g = (0, react_1.useState)(''), tab3Error = _g[0], setTab3Error = _g[1];
    var _h = (0, react_1.useState)(''), removingUrl = _h[0], setRemovingUrl = _h[1];
    var abortControllerRef = (0, react_1.useRef)(null);
    var patchState = (0, react_1.useCallback)(function (patch) {
        setState(function (prev) { return (tslib_1.__assign(tslib_1.__assign({}, prev), patch)); });
    }, []);
    // ── Tab 1: load current user + saved tags from userPersonalization ───────────
    var loadTab1Data = (0, react_1.useCallback)(function () { return tslib_1.__awaiter(void 0, void 0, void 0, function () {
        var currentUser, record, err_1, msg;
        var _a, _b, _c;
        return tslib_1.__generator(this, function (_d) {
            switch (_d.label) {
                case 0:
                    patchState({ isLoadingTags: true, tab1Error: '', tab1Success: '' });
                    _d.label = 1;
                case 1:
                    _d.trys.push([1, 4, , 5]);
                    return [4 /*yield*/, spService.current.getCurrentUser()];
                case 2:
                    currentUser = _d.sent();
                    return [4 /*yield*/, spService.current.getUserPersonalizationByUserId(userPersonalizationListName, currentUser.Id)];
                case 3:
                    record = _d.sent();
                    patchState({
                        savedTags: (_a = record === null || record === void 0 ? void 0 : record.SelectedTags) !== null && _a !== void 0 ? _a : '',
                        isLoadingTags: false,
                        currentUserId: currentUser.Id,
                        currentUserLoginName: currentUser.LoginName,
                        userPersonalizationItemId: (_b = record === null || record === void 0 ? void 0 : record.itemId) !== null && _b !== void 0 ? _b : null,
                        savedLinks: (_c = record === null || record === void 0 ? void 0 : record.SavedLinks) !== null && _c !== void 0 ? _c : ''
                    });
                    return [3 /*break*/, 5];
                case 4:
                    err_1 = _d.sent();
                    msg = err_1 instanceof Error ? err_1.message : String(err_1);
                    patchState({ isLoadingTags: false, tab1Error: msg });
                    return [3 /*break*/, 5];
                case 5: return [2 /*return*/];
            }
        });
    }); }, [userPersonalizationListName, patchState]);
    // ── Tab 2: fetch article recommendations for current user ──────────────────
    var loadTab2Data = (0, react_1.useCallback)(function () { return tslib_1.__awaiter(void 0, void 0, void 0, function () {
        var controller, currentUser, record, tags, articles, _a;
        var _b, _c, _d;
        return tslib_1.__generator(this, function (_e) {
            switch (_e.label) {
                case 0:
                    if (abortControllerRef.current) {
                        abortControllerRef.current.abort();
                    }
                    controller = new AbortController();
                    abortControllerRef.current = controller;
                    patchState({ isLoadingArticles: true, tab2Error: '', tab2Info: '', articles: [] });
                    _e.label = 1;
                case 1:
                    _e.trys.push([1, 5, , 6]);
                    return [4 /*yield*/, spService.current.getCurrentUser()];
                case 2:
                    currentUser = _e.sent();
                    return [4 /*yield*/, spService.current.getUserPersonalizationByUserId(userPersonalizationListName, currentUser.Id)];
                case 3:
                    record = _e.sent();
                    if (!record) {
                        if (!controller.signal.aborted) {
                            patchState({
                                isLoadingArticles: false,
                                tab2Info: 'No interests saved yet. Go to the My Interests tab to select your topics.'
                            });
                        }
                        return [2 /*return*/];
                    }
                    tags = (_c = (_b = record.SelectedTags) === null || _b === void 0 ? void 0 : _b.trim()) !== null && _c !== void 0 ? _c : '';
                    if (!tags) {
                        if (!controller.signal.aborted) {
                            patchState({
                                isLoadingArticles: false,
                                tab2Info: 'Your interests list is empty. Please select topics in the My Interests tab.'
                            });
                        }
                        return [2 /*return*/];
                    }
                    patchState({
                        userPersonalizationItemId: record.itemId,
                        currentUserId: currentUser.Id,
                        currentUserLoginName: currentUser.LoginName,
                        savedLinks: (_d = record.SavedLinks) !== null && _d !== void 0 ? _d : ''
                    });
                    return [4 /*yield*/, topicsServiceRef.current.getArticles(tags, 20)];
                case 4:
                    articles = _e.sent();
                    if (!controller.signal.aborted) {
                        if (articles.length === 0) {
                            patchState({
                                isLoadingArticles: false,
                                tab2Info: 'No articles found for your selected topics. Try updating your interests.'
                            });
                        }
                        else {
                            patchState({ isLoadingArticles: false, articles: articles });
                        }
                    }
                    return [3 /*break*/, 6];
                case 5:
                    _a = _e.sent();
                    if (!controller.signal.aborted) {
                        patchState({
                            isLoadingArticles: false,
                            tab2Error: 'Unable to fetch articles. Please try again later.'
                        });
                    }
                    return [3 /*break*/, 6];
                case 6: return [2 /*return*/];
            }
        });
    }); }, [userPersonalizationListName, patchState]);
    // ── Tab 3: refresh SavedLinks from userPersonalization ────────────────────
    var refreshSavedLinks = (0, react_1.useCallback)(function () { return tslib_1.__awaiter(void 0, void 0, void 0, function () {
        var userId, user, record, err_2;
        var _a, _b;
        return tslib_1.__generator(this, function (_c) {
            switch (_c.label) {
                case 0:
                    setIsLoadingSavedLinks(true);
                    setTab3Error('');
                    _c.label = 1;
                case 1:
                    _c.trys.push([1, 5, 6, 7]);
                    userId = state.currentUserId;
                    if (!!userId) return [3 /*break*/, 3];
                    return [4 /*yield*/, spService.current.getCurrentUser()];
                case 2:
                    user = _c.sent();
                    userId = user.Id;
                    patchState({ currentUserId: user.Id, currentUserLoginName: user.LoginName });
                    _c.label = 3;
                case 3: return [4 /*yield*/, spService.current.getUserPersonalizationByUserId(userPersonalizationListName, userId)];
                case 4:
                    record = _c.sent();
                    patchState({
                        savedLinks: (_a = record === null || record === void 0 ? void 0 : record.SavedLinks) !== null && _a !== void 0 ? _a : '',
                        userPersonalizationItemId: (_b = record === null || record === void 0 ? void 0 : record.itemId) !== null && _b !== void 0 ? _b : null
                    });
                    return [3 /*break*/, 7];
                case 5:
                    err_2 = _c.sent();
                    setTab3Error(err_2 instanceof Error ? err_2.message : String(err_2));
                    return [3 /*break*/, 7];
                case 6:
                    setIsLoadingSavedLinks(false);
                    return [7 /*endfinally*/];
                case 7: return [2 /*return*/];
            }
        });
    }); }, [state.currentUserId, userPersonalizationListName, patchState]);
    // ── Mount: load Tab 1 data ─────────────────────────────────────────────────
    (0, react_1.useEffect)(function () {
        loadTab2Data().catch(function (err) {
            return console.error('AI Curator: Failed to load Tab 2 data', err);
        });
        return function () {
            if (abortControllerRef.current) {
                abortControllerRef.current.abort();
            }
        };
    }, []);
    // ── Tab switch ─────────────────────────────────────────────────────────────
    var handleTabChange = (0, react_1.useCallback)(function (item) {
        var _a;
        if (!item)
            return;
        var key = (_a = item.props.itemKey) !== null && _a !== void 0 ? _a : TAB_ARTICLES;
        patchState({ activeTab: key });
        if (key === TAB_ARTICLES) {
            loadTab2Data().catch(function (err) {
                return console.error('AI Curator: Failed to load Tab 2 data', err);
            });
        }
        else if (key === TAB_SAVED) {
            refreshSavedLinks().catch(function (err) {
                return console.error('AI Curator: Failed to refresh saved links', err);
            });
        }
        else if (key === TAB_INTERESTS) {
            loadTab1Data().catch(function (err) {
                return console.error('AI Curator: Failed to load Tab 1 data', err);
            });
        }
    }, [loadTab1Data, loadTab2Data, refreshSavedLinks, patchState]);
    // ── Save Interests (upsert userPersonalization) ────────────────────────────
    var handleSaveInterests = (0, react_1.useCallback)(function (selectedTags) { return tslib_1.__awaiter(void 0, void 0, void 0, function () {
        var userId, loginName, user, tagsString, created, err_3, msg;
        return tslib_1.__generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    setIsSavingTags(true);
                    patchState({ tab1Error: '', tab1Success: '' });
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 8, 9, 10]);
                    userId = state.currentUserId;
                    loginName = state.currentUserLoginName;
                    if (!!userId) return [3 /*break*/, 3];
                    return [4 /*yield*/, spService.current.getCurrentUser()];
                case 2:
                    user = _a.sent();
                    userId = user.Id;
                    loginName = user.LoginName;
                    patchState({ currentUserId: userId, currentUserLoginName: loginName });
                    _a.label = 3;
                case 3:
                    tagsString = selectedTags.join(', ');
                    if (!state.userPersonalizationItemId) return [3 /*break*/, 5];
                    return [4 /*yield*/, spService.current.updateSelectedTags(userPersonalizationListName, state.userPersonalizationItemId, tagsString)];
                case 4:
                    _a.sent();
                    return [3 /*break*/, 7];
                case 5: return [4 /*yield*/, spService.current.createUserPersonalization(userPersonalizationListName, loginName, userId, tagsString)];
                case 6:
                    created = _a.sent();
                    patchState({ userPersonalizationItemId: created.itemId });
                    _a.label = 7;
                case 7:
                    patchState({ tab1Success: 'Your interests have been saved successfully!', savedTags: tagsString });
                    setTimeout(function () {
                        patchState({ activeTab: TAB_ARTICLES, tab1Success: '' });
                        loadTab2Data().catch(function (err) {
                            return console.error('AI Curator: Failed to load Tab 2 after save', err);
                        });
                    }, 1200);
                    return [3 /*break*/, 10];
                case 8:
                    err_3 = _a.sent();
                    msg = err_3 instanceof Error ? err_3.message : String(err_3);
                    patchState({ tab1Error: "Failed to save interests: ".concat(msg) });
                    return [3 /*break*/, 10];
                case 9:
                    setIsSavingTags(false);
                    return [7 /*endfinally*/];
                case 10: return [2 /*return*/];
            }
        });
    }); }, [state.currentUserId, state.currentUserLoginName, state.userPersonalizationItemId, userPersonalizationListName, loadTab2Data, patchState]);
    // ── Save article link ──────────────────────────────────────────────────────
    var handleSaveArticle = (0, react_1.useCallback)(function (article) { return tslib_1.__awaiter(void 0, void 0, void 0, function () {
        var itemId, updated;
        return tslib_1.__generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    itemId = state.userPersonalizationItemId;
                    if (!itemId) {
                        throw new Error('Could not save — user personalization record not found. ' +
                            'Please save your interests first in the My Interests tab.');
                    }
                    return [4 /*yield*/, spService.current.saveArticleLink(userPersonalizationListName, itemId, state.savedLinks, article.url)];
                case 1:
                    _a.sent();
                    updated = state.savedLinks
                        ? "".concat(state.savedLinks, ",").concat(article.url)
                        : article.url;
                    patchState({ savedLinks: updated });
                    return [2 /*return*/];
            }
        });
    }); }, [state.userPersonalizationItemId, state.savedLinks, userPersonalizationListName, patchState]);
    // ── Remove saved link ──────────────────────────────────────────────────────
    var handleRemoveSavedLink = (0, react_1.useCallback)(function (url) { return tslib_1.__awaiter(void 0, void 0, void 0, function () {
        var itemId, updated, err_4;
        return tslib_1.__generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    itemId = state.userPersonalizationItemId;
                    if (!itemId)
                        return [2 /*return*/];
                    setRemovingUrl(url);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, 4, 5]);
                    return [4 /*yield*/, spService.current.removeSavedLink(userPersonalizationListName, itemId, state.savedLinks, url)];
                case 2:
                    _a.sent();
                    updated = state.savedLinks
                        .split(',')
                        .map(function (l) { return l.trim(); })
                        .filter(function (l) { return l.length > 0 && l !== url; })
                        .join(',');
                    patchState({ savedLinks: updated });
                    return [3 /*break*/, 5];
                case 3:
                    err_4 = _a.sent();
                    setTab3Error(err_4 instanceof Error ? err_4.message : String(err_4));
                    return [3 /*break*/, 5];
                case 4:
                    setRemovingUrl('');
                    return [7 /*endfinally*/];
                case 5: return [2 /*return*/];
            }
        });
    }); }, [state.userPersonalizationItemId, state.savedLinks, userPersonalizationListName, patchState]);
    // ── Share panel ────────────────────────────────────────────────────────────
    var loadYammerGroups = (0, react_1.useCallback)(function () {
        setIsLoadingGroups(true);
        setYammerGroupsError('');
        webPartContext.msGraphClientFactory
            .getClient('3')
            .then(function (graphClient) { return new VivaEngageService_1.VivaEngageService(graphClient).getYammerGroups(); })
            .then(function (groups) {
            setYammerGroups(groups);
            setIsLoadingGroups(false);
        })
            .catch(function (err) {
            var msg = err instanceof Error ? err.message : String(err);
            setYammerGroupsError(msg);
            setIsLoadingGroups(false);
        });
    }, [webPartContext]);
    var handleShareArticle = (0, react_1.useCallback)(function (article) {
        var _a;
        patchState({
            sharePanelArticleUrl: article.url,
            sharePanelArticleTitle: article.title,
            sharePanelArticleSummary: (_a = article.summary) !== null && _a !== void 0 ? _a : ''
        });
        // Always (re-)fetch groups when the panel opens so stale/failed state is cleared
        loadYammerGroups();
    }, [loadYammerGroups, patchState]);
    var handleLinkedInShareArticle = (0, react_1.useCallback)(function (article) {
        var _a;
        patchState({
            linkedInPanelArticleUrl: article.url,
            linkedInPanelArticleTitle: article.title,
            linkedInPanelArticleSummary: (_a = article.summary) !== null && _a !== void 0 ? _a : ''
        });
    }, [patchState]);
    var handlePostToVivaEngage = (0, react_1.useCallback)(function (groupId, userComments) { return tslib_1.__awaiter(void 0, void 0, void 0, function () {
        var graphClient;
        return tslib_1.__generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4 /*yield*/, webPartContext.msGraphClientFactory.getClient('3')];
                case 1:
                    graphClient = _a.sent();
                    return [4 /*yield*/, new VivaEngageService_1.VivaEngageService(graphClient).postToGroup(groupId, state.sharePanelArticleUrl, userComments)];
                case 2:
                    _a.sent();
                    return [2 /*return*/];
            }
        });
    }); }, [webPartContext, state.sharePanelArticleUrl]);
    // ── Derived ────────────────────────────────────────────────────────────────
    var savedArticleUrls = state.savedLinks
        .split(',')
        .map(function (l) { return l.trim(); })
        .filter(function (l) { return l.length > 0; });
    var sharePanelIsOpen = state.sharePanelArticleUrl.length > 0 && state.sharePanelArticleTitle.length > 0;
    var linkedInPanelIsOpen = state.linkedInPanelArticleUrl.length > 0 && state.linkedInPanelArticleTitle.length > 0;
    // ── Render ─────────────────────────────────────────────────────────────────
    return (React.createElement("section", { className: "".concat(AiCuratorArticleRecommender_module_scss_1.default.aiCuratorArticleRecommender, " ").concat(hasTeamsContext ? AiCuratorArticleRecommender_module_scss_1.default.teams : '', " ").concat(isDarkTheme ? AiCuratorArticleRecommender_module_scss_1.default.darkTheme : '') },
        React.createElement(react_2.Stack, { tokens: { childrenGap: 16 }, className: AiCuratorArticleRecommender_module_scss_1.default.container },
            React.createElement(react_2.Stack, { tokens: { childrenGap: 4 }, className: AiCuratorArticleRecommender_module_scss_1.default.header },
                React.createElement(react_2.Stack, { horizontal: true, verticalAlign: "center", tokens: { childrenGap: 8 } },
                    React.createElement(react_2.Icon, { iconName: "LightningBolt", className: AiCuratorArticleRecommender_module_scss_1.default.headerIcon }),
                    React.createElement(react_2.Text, { variant: "xLarge", className: AiCuratorArticleRecommender_module_scss_1.default.title }, "AI Curator \u2013 Article Recommender"))),
            React.createElement(react_2.Separator, null),
            React.createElement(react_2.Pivot, { selectedKey: state.activeTab, onLinkClick: handleTabChange, styles: {
                    linkIsSelected: {
                        color: '#107C10',
                        selectors: { '::before': { backgroundColor: '#107C10' } }
                    }
                } },
                React.createElement(react_2.PivotItem, { headerText: "Recommended Articles", itemKey: TAB_ARTICLES, itemIcon: "Lightbulb" },
                    React.createElement(ArticleList_1.default, { articles: state.articles, isLoading: state.isLoadingArticles, errorMessage: state.tab2Error, infoMessage: state.tab2Info, savedArticleUrls: savedArticleUrls, onSaveArticle: handleSaveArticle, onShareArticle: handleShareArticle, onLinkedInShareArticle: handleLinkedInShareArticle, vivaEngageEnabled: vivaEngageEnabled })),
                React.createElement(react_2.PivotItem, { headerText: "My Saved Links", itemKey: TAB_SAVED, itemIcon: "FavoriteStar" },
                    tab3Error && (React.createElement(react_2.MessageBar, { messageBarType: react_2.MessageBarType.error, isMultiline: true, onDismiss: function () { return setTab3Error(''); }, style: { marginTop: 8, marginBottom: 4 } }, tab3Error)),
                    isLoadingSavedLinks ? (React.createElement(react_2.Stack, { horizontalAlign: "center", style: { padding: 40 } },
                        React.createElement(react_2.Spinner, { size: react_2.SpinnerSize.large, label: "Loading saved links\u2026", labelPosition: "bottom" }))) : savedArticleUrls.length === 0 ? (React.createElement(react_2.Stack, { horizontalAlign: "center", tokens: { childrenGap: 8 }, style: { padding: '40px 16px' } },
                        React.createElement(react_2.Icon, { iconName: "FavoriteStar", style: { fontSize: 36, color: '#c8c6c4' } }),
                        React.createElement(react_2.Text, { variant: "mediumPlus", style: { color: '#605e5c', fontWeight: 600 } }, "No saved links yet"),
                        React.createElement(react_2.Text, { variant: "small", style: { color: '#a19f9d', textAlign: 'center' } }, "Articles you save from the Recommended Articles tab will appear here."))) : (React.createElement(react_2.Stack, { tokens: { childrenGap: 8 }, style: { marginTop: 12 } },
                        React.createElement(react_2.Text, { variant: "small", style: { color: '#605e5c', marginBottom: 4 } },
                            savedArticleUrls.length,
                            " saved ",
                            savedArticleUrls.length === 1 ? 'link' : 'links'),
                        savedArticleUrls.map(function (url, i) { return (React.createElement(react_2.Stack, { key: "saved-".concat(i, "-").concat(url), horizontal: true, verticalAlign: "center", tokens: { childrenGap: 10 }, style: {
                                padding: '10px 14px',
                                borderRadius: '8px',
                                backgroundColor: '#ffffff',
                                border: '1px solid #edebe9',
                                boxShadow: '0 1px 3px rgba(0,0,0,0.08)'
                            } },
                            React.createElement(react_2.Icon, { iconName: "Link", style: { fontSize: 16, color: '#107C10', flexShrink: 0 } }),
                            React.createElement(react_2.Link, { href: url, target: "_blank", rel: "noopener noreferrer", style: { flexGrow: 1, wordBreak: 'break-all', color: '#107C10', fontSize: 13 } }, url),
                            React.createElement(react_2.TooltipHost, { content: "Remove link" },
                                React.createElement(react_2.IconButton, { iconProps: { iconName: 'Delete' }, ariaLabel: "Remove saved link", disabled: removingUrl === url, onClick: function () { void handleRemoveSavedLink(url); }, styles: {
                                        root: { color: '#605e5c', flexShrink: 0 },
                                        rootHovered: { color: '#a4262c' },
                                        icon: { fontSize: 14 }
                                    } })))); })))),
                React.createElement(react_2.PivotItem, { headerText: "My Interests", itemKey: TAB_INTERESTS, itemIcon: "Tag" },
                    state.tab1Error && (React.createElement(react_2.MessageBar, { messageBarType: react_2.MessageBarType.error, isMultiline: true, onDismiss: function () { return patchState({ tab1Error: '' }); }, style: { marginTop: 8, marginBottom: 4 } }, state.tab1Error)),
                    state.isLoadingTags ? (React.createElement(react_2.Stack, { horizontalAlign: "center", style: { padding: 40 } },
                        React.createElement(react_2.Spinner, { size: react_2.SpinnerSize.large, label: "Loading your interests\u2026", labelPosition: "bottom" }))) : (React.createElement(TagSelector_1.default, { savedTags: state.savedTags, successMessage: state.tab1Success, onSave: function (tags) { void handleSaveInterests(tags); }, isSaving: isSavingTags }))))),
        React.createElement(SharePanel_1.default, { isOpen: sharePanelIsOpen, articleUrl: state.sharePanelArticleUrl, articleTitle: state.sharePanelArticleTitle, articleSummary: state.sharePanelArticleSummary, groups: yammerGroups, isLoadingGroups: isLoadingGroups, groupLoadError: yammerGroupsError, onRetryLoadGroups: loadYammerGroups, onDismiss: function () { return patchState({ sharePanelArticleUrl: '', sharePanelArticleTitle: '', sharePanelArticleSummary: '' }); }, onPost: handlePostToVivaEngage }),
        React.createElement(LinkedInPanel_1.default, { isOpen: linkedInPanelIsOpen, articleUrl: state.linkedInPanelArticleUrl, articleTitle: state.linkedInPanelArticleTitle, articleSummary: state.linkedInPanelArticleSummary, onDismiss: function () { return patchState({ linkedInPanelArticleUrl: '', linkedInPanelArticleTitle: '', linkedInPanelArticleSummary: '' }); } })));
};
exports.default = AiCuratorArticleRecommender;
//# sourceMappingURL=AiCuratorArticleRecommender.js.map