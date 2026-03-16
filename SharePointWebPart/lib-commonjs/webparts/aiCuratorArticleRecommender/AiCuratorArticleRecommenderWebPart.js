"use strict";
/**
 * ============================================================================
 * AI Curator - Article Recommender Web Part
 * ============================================================================
 *
 * *** HEFT-BASED BUILD (SPFx v1.22.1+) ***
 * Build rig: @microsoft/spfx-web-build-rig
 *
 * DEPLOYMENT (Heft-based):
 *  1. npm install
 *  2. npm run build              (heft build --clean)
 *  3. npm run bundle             (heft bundle --production)
 *  4. npm run package-solution   (heft package-solution --production)
 *  5. Upload the .sppkg from sharepoint/solution/ to your App Catalog
 *
 * TOPICS DATA FLOW:
 *  Topics are discovered via the external suggest-topics API.
 *  Article recommendations come from the external articles API.
 *  No OpenAI keys or endpoints are stored in this web part.
 *
 * ============================================================================
 */
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var React = tslib_1.__importStar(require("react"));
var ReactDom = tslib_1.__importStar(require("react-dom"));
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_property_pane_1 = require("@microsoft/sp-property-pane");
var sp_webpart_base_1 = require("@microsoft/sp-webpart-base");
var AiCuratorArticleRecommender_1 = tslib_1.__importDefault(require("./components/AiCuratorArticleRecommender"));
var strings = tslib_1.__importStar(require("AiCuratorArticleRecommenderWebPartStrings"));
var AiCuratorArticleRecommenderWebPart = /** @class */ (function (_super) {
    tslib_1.__extends(AiCuratorArticleRecommenderWebPart, _super);
    function AiCuratorArticleRecommenderWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this._isDarkTheme = false;
        return _this;
    }
    AiCuratorArticleRecommenderWebPart.prototype.onInit = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, _super.prototype.onInit.call(this)];
                    case 1:
                        _a.sent();
                        return [2 /*return*/];
                }
            });
        });
    };
    AiCuratorArticleRecommenderWebPart.prototype.render = function () {
        var _a, _b;
        var element = React.createElement(AiCuratorArticleRecommender_1.default, {
            articlesListName: this.properties.articlesListName || 'Articles',
            userPersonalizationListName: this.properties.userPersonalizationListName || 'userPersonalization',
            articlesLimit: (_a = this.properties.articlesLimit) !== null && _a !== void 0 ? _a : 10,
            isDarkTheme: this._isDarkTheme,
            hasTeamsContext: !!((_b = this.context.sdks) === null || _b === void 0 ? void 0 : _b.microsoftTeams),
            webPartContext: this.context
        });
        ReactDom.render(element, this.domElement);
    };
    AiCuratorArticleRecommenderWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(AiCuratorArticleRecommenderWebPart.prototype, "dataVersion", {
        get: function () {
            return sp_core_library_1.Version.parse('1.0');
        },
        enumerable: false,
        configurable: true
    });
    Object.defineProperty(AiCuratorArticleRecommenderWebPart.prototype, "disableReactivePropertyChanges", {
        get: function () {
            return false;
        },
        enumerable: false,
        configurable: true
    });
    /**
     * Property pane — two groups.
     * NO keyword, tag, or OpenAI fields anywhere in this pane.
     * Topics are discovered via external API at runtime.
     */
    AiCuratorArticleRecommenderWebPart.prototype.getPropertyPaneConfiguration = function () {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription
                    },
                    groups: [
                        // ── Group 1: SharePoint Data Source ───────────────────────────
                        {
                            groupName: strings.DataSourceGroupName,
                            groupFields: [
                                (0, sp_property_pane_1.PropertyPaneTextField)('articlesListName', {
                                    label: strings.ListNameFieldLabel,
                                    placeholder: 'Articles',
                                    description: 'Display name of the Articles list (used for context only).',
                                    value: 'Articles'
                                })
                            ]
                        },
                        // ── Group 2: Personalization & Sharing ────────────────────────
                        {
                            groupName: strings.PersonalizationGroupName,
                            groupFields: [
                                (0, sp_property_pane_1.PropertyPaneTextField)('userPersonalizationListName', {
                                    label: strings.UserPersonalizationListNameFieldLabel,
                                    description: 'Display name of the list storing user tag selections and saved links.',
                                    value: 'userPersonalization'
                                }),
                                (0, sp_property_pane_1.PropertyPaneSlider)('articlesLimit', {
                                    label: strings.ArticlesLimitFieldLabel,
                                    min: 1,
                                    max: 100,
                                    step: 1,
                                    showValue: true
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return AiCuratorArticleRecommenderWebPart;
}(sp_webpart_base_1.BaseClientSideWebPart));
exports.default = AiCuratorArticleRecommenderWebPart;
//# sourceMappingURL=AiCuratorArticleRecommenderWebPart.js.map