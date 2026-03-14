"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var React = tslib_1.__importStar(require("react"));
var react_1 = require("react");
var react_2 = require("@fluentui/react");
var GREEN = '#107C10';
/** Format an ISO date string to a readable short date, e.g. "13 Mar 2026" */
function formatPublished(iso) {
    try {
        return new Date(iso).toLocaleDateString(undefined, {
            day: 'numeric',
            month: 'short',
            year: 'numeric'
        });
    }
    catch (_a) {
        return iso;
    }
}
/**
 * A single article card with green-themed styling, title link,
 * description, URL, and Save / Share action buttons.
 */
var ArticleCard = function (props) {
    var article = props.article, index = props.index, isSaved = props.isSaved, onSave = props.onSave, onShare = props.onShare;
    var _a = (0, react_1.useState)(false), isSaving = _a[0], setIsSaving = _a[1];
    var _b = (0, react_1.useState)(isSaved), saved = _b[0], setSaved = _b[1];
    var handleSave = function () { return tslib_1.__awaiter(void 0, void 0, void 0, function () {
        return tslib_1.__generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (saved || isSaving)
                        return [2 /*return*/];
                    setIsSaving(true);
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, , 3, 4]);
                    return [4 /*yield*/, onSave(article)];
                case 2:
                    _a.sent();
                    setSaved(true);
                    return [3 /*break*/, 4];
                case 3:
                    setIsSaving(false);
                    return [7 /*endfinally*/];
                case 4: return [2 /*return*/];
            }
        });
    }); };
    var displayUrl = article.url.length > 80
        ? article.url.substring(0, 77) + '…'
        : article.url;
    return (React.createElement(react_2.Stack, { key: "article-".concat(index, "-").concat(article.url), style: {
            padding: '12px 16px',
            borderRadius: '8px',
            boxShadow: '0 1px 3px rgba(0,0,0,0.12), 0 1px 2px rgba(0,0,0,0.06)',
            transition: 'box-shadow 0.2s ease, transform 0.15s ease',
            backgroundColor: '#ffffff',
            border: "1px solid #edebe9",
            marginBottom: 0
        } },
        React.createElement(react_2.Stack, { horizontal: true, verticalAlign: "start", tokens: { childrenGap: 10 } },
            React.createElement(react_2.Icon, { iconName: "TextDocument", style: { fontSize: 20, color: GREEN, marginTop: 2, flexShrink: 0 } }),
            React.createElement(react_2.Stack, { tokens: { childrenGap: 4 }, grow: true },
                React.createElement(react_2.Link, { href: article.url, target: "_blank", rel: "noopener noreferrer", style: {
                        fontSize: 15,
                        fontWeight: 600,
                        lineHeight: '1.3',
                        color: GREEN,
                        textDecoration: 'none'
                    }, title: "Open \"".concat(article.title, "\" in a new tab") }, article.title),
                article.summary && (React.createElement(react_2.Text, { variant: "small", style: { color: '#323130', lineHeight: '1.6', whiteSpace: 'pre-line' } }, article.summary)),
                React.createElement(react_2.Stack, { horizontal: true, tokens: { childrenGap: 12 }, style: { flexWrap: 'wrap' } },
                    article.source && (React.createElement(react_2.Stack, { horizontal: true, verticalAlign: "center", tokens: { childrenGap: 4 } },
                        React.createElement(react_2.Icon, { iconName: "Globe", style: { fontSize: 12, color: '#605e5c' } }),
                        React.createElement(react_2.Text, { variant: "tiny", style: { color: '#605e5c', fontWeight: 600 } }, article.source))),
                    article.published && (React.createElement(react_2.Stack, { horizontal: true, verticalAlign: "center", tokens: { childrenGap: 4 } },
                        React.createElement(react_2.Icon, { iconName: "Calendar", style: { fontSize: 12, color: '#605e5c' } }),
                        React.createElement(react_2.Text, { variant: "tiny", style: { color: '#605e5c' } }, formatPublished(article.published))))),
                React.createElement(react_2.Text, { variant: "tiny", style: { color: '#605e5c', wordBreak: 'break-all', opacity: 0.7 } }, displayUrl)),
            React.createElement(react_2.Stack, { horizontal: true, tokens: { childrenGap: 4 }, style: { flexShrink: 0, alignSelf: 'flex-start' } },
                React.createElement(react_2.TooltipHost, { content: saved ? 'Already saved' : 'Save to My Interests' },
                    React.createElement(react_2.IconButton, { iconProps: { iconName: saved ? 'CheckMark' : 'Save' }, ariaLabel: saved ? 'Already saved' : 'Save article', disabled: saved || isSaving, onClick: function () { void handleSave(); }, styles: {
                            root: { color: saved ? GREEN : '#605e5c' },
                            rootDisabled: { color: GREEN },
                            icon: { fontSize: 16 }
                        } })),
                React.createElement(react_2.TooltipHost, { content: "Share to Viva Engage" },
                    React.createElement(react_2.IconButton, { iconProps: { iconName: 'Share' }, ariaLabel: "Share to Viva Engage", onClick: function () { return onShare(article); }, styles: {
                            root: { color: '#605e5c' },
                            rootHovered: { color: GREEN },
                            icon: { fontSize: 16 }
                        } }))))));
};
exports.default = ArticleCard;
//# sourceMappingURL=ArticleCard.js.map