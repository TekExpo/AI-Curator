"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var React = tslib_1.__importStar(require("react"));
var react_1 = require("@fluentui/react");
var ArticleCard_1 = tslib_1.__importDefault(require("./ArticleCard"));
/**
 * Tab 2 – Recommended Articles
 * Renders the AI-powered article recommendation list.
 */
var ArticleList = function (props) {
    var articles = props.articles, isLoading = props.isLoading, errorMessage = props.errorMessage, infoMessage = props.infoMessage, savedArticleUrls = props.savedArticleUrls, onSaveArticle = props.onSaveArticle, onShareArticle = props.onShareArticle, vivaEngageEnabled = props.vivaEngageEnabled;
    if (isLoading) {
        return (React.createElement(react_1.Stack, { horizontalAlign: "center", verticalAlign: "center", style: { minHeight: 200, padding: 40 } },
            React.createElement(react_1.Spinner, { size: react_1.SpinnerSize.large, label: "Fetching article recommendations\u2026", labelPosition: "bottom" })));
    }
    return (React.createElement(react_1.Stack, { tokens: { childrenGap: 12 }, style: { padding: '8px 0' } },
        errorMessage && (React.createElement(react_1.MessageBar, { messageBarType: react_1.MessageBarType.error, isMultiline: true }, errorMessage)),
        infoMessage && !errorMessage && (React.createElement(react_1.MessageBar, { messageBarType: react_1.MessageBarType.info, isMultiline: true }, infoMessage)),
        !isLoading && !errorMessage && !infoMessage && articles.length === 0 && (React.createElement(react_1.MessageBar, { messageBarType: react_1.MessageBarType.info }, "No article recommendations were returned. Try updating your interests in the My Interests tab.")),
        articles.map(function (article, index) { return (React.createElement(ArticleCard_1.default, { key: "article-".concat(index, "-").concat(article.url), article: article, index: index, isSaved: savedArticleUrls.indexOf(article.url) !== -1, vivaEngageEnabled: vivaEngageEnabled, onSave: onSaveArticle, onShare: onShareArticle })); }),
        articles.length > 0 && (React.createElement(react_1.Text, { variant: "tiny", style: {
                textAlign: 'center',
                color: '#605e5c',
                paddingTop: 8,
                opacity: 0.6
            } },
            "Showing ",
            articles.length,
            " recommendation",
            articles.length !== 1 ? 's' : '',
            " \u2022 Powered by AI"))));
};
exports.default = ArticleList;
//# sourceMappingURL=ArticleList.js.map