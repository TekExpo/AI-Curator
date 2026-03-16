"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var React = tslib_1.__importStar(require("react"));
var react_1 = require("react");
// Load Quill's snow theme CSS via webpack CSS loader (SPFx-compatible pattern)
// eslint-disable-next-line @typescript-eslint/no-require-imports
require('quill/dist/quill.snow.css');
var react_2 = require("@fluentui/react");
var react_quill_1 = tslib_1.__importDefault(require("react-quill"));
var GREEN = '#107C10';
/** Quill toolbar configuration – basic rich-text set */
var QUILL_MODULES = {
    toolbar: [
        ['bold', 'italic', 'underline'],
        [{ list: 'ordered' }, { list: 'bullet' }],
        ['link'],
        ['clean']
    ]
};
var QUILL_FORMATS = ['bold', 'italic', 'underline', 'list', 'bullet', 'link'];
/** Escape plain text for safe inclusion in HTML */
var escapeHtml = function (text) {
    return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
};
/** Convert a plain-text summary (with newlines) to basic HTML paragraphs */
var summaryToHtml = function (text) {
    if (!text)
        return '';
    return text
        .split('\n')
        .map(function (line) { return "<p>".concat(escapeHtml(line) || '<br>', "</p>"); })
        .join('');
};
/** Strip HTML tags and return plain text (for clipboard copy) */
var htmlToPlainText = function (html) {
    return html.replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').trim();
};
/**
 * Slide-in panel for sharing an article to a Viva Engage group.
 */
var SharePanel = function (props) {
    var isOpen = props.isOpen, articleUrl = props.articleUrl, articleTitle = props.articleTitle, articleSummary = props.articleSummary, groups = props.groups, isLoadingGroups = props.isLoadingGroups, groupLoadError = props.groupLoadError, onRetryLoadGroups = props.onRetryLoadGroups, onDismiss = props.onDismiss, onPost = props.onPost;
    var _a = (0, react_1.useState)(''), selectedGroupId = _a[0], setSelectedGroupId = _a[1];
    var _b = (0, react_1.useState)(''), description = _b[0], setDescription = _b[1];
    var _c = (0, react_1.useState)(false), isPosting = _c[0], setIsPosting = _c[1];
    var _d = (0, react_1.useState)(''), errorMessage = _d[0], setErrorMessage = _d[1];
    var _f = (0, react_1.useState)(''), successMessage = _f[0], setSuccessMessage = _f[1];
    var _g = (0, react_1.useState)(false), copied = _g[0], setCopied = _g[1];
    // Reset and pre-populate whenever the panel opens
    (0, react_1.useEffect)(function () {
        if (isOpen) {
            setSelectedGroupId('');
            setDescription((articleSummary === null || articleSummary === void 0 ? void 0 : articleSummary.trim()) ? summaryToHtml(articleSummary.trim()) : '');
            setErrorMessage('');
            setSuccessMessage('');
            setIsPosting(false);
            setCopied(false);
        }
    }, [isOpen]);
    var handleDismiss = function () {
        setSelectedGroupId('');
        setDescription('');
        setErrorMessage('');
        setSuccessMessage('');
        setCopied(false);
        onDismiss();
    };
    var handleCopyToClipboard = function () {
        var plainText = [
            htmlToPlainText(description),
            articleUrl,
            'Shared via AI Curator \u2013 Article Recommender'
        ].filter(Boolean).join('\n\n');
        if (navigator.clipboard && plainText) {
            navigator.clipboard.writeText(plainText).then(function () {
                setCopied(true);
                setTimeout(function () { return setCopied(false); }, 2000);
            }).catch(function () { });
        }
    };
    var handlePost = function () { return tslib_1.__awaiter(void 0, void 0, void 0, function () {
        var descText, err_1, msg;
        return tslib_1.__generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (!selectedGroupId) {
                        setErrorMessage('Please select a Viva Engage community.');
                        return [2 /*return*/];
                    }
                    descText = description.replace(/<[^>]*>/g, '').trim();
                    if (!descText) {
                        setErrorMessage('Please enter a description for the post.');
                        return [2 /*return*/];
                    }
                    setIsPosting(true);
                    setErrorMessage('');
                    setSuccessMessage('');
                    _a.label = 1;
                case 1:
                    _a.trys.push([1, 3, 4, 5]);
                    return [4 /*yield*/, onPost(selectedGroupId, description)];
                case 2:
                    _a.sent();
                    setSuccessMessage('Article successfully shared to Viva Engage!');
                    return [3 /*break*/, 5];
                case 3:
                    err_1 = _a.sent();
                    msg = err_1 instanceof Error ? err_1.message : String(err_1);
                    setErrorMessage(msg);
                    return [3 /*break*/, 5];
                case 4:
                    setIsPosting(false);
                    return [7 /*endfinally*/];
                case 5: return [2 /*return*/];
            }
        });
    }); };
    var groupOptions = groups.map(function (g) { return ({
        key: g.id,
        text: g.name
    }); });
    var previewHtml = (description.trim() || '') +
        (articleUrl
            ? "<p><a href=\"".concat(escapeHtml(articleUrl), "\" target=\"_blank\" rel=\"noopener noreferrer\">").concat(escapeHtml(articleUrl), "</a></p>")
            : '') +
        '<p><em>Shared via AI Curator \u2013 Article Recommender</em></p>';
    return (React.createElement(react_2.Panel, { isOpen: isOpen, onDismiss: handleDismiss, type: react_2.PanelType.medium, headerText: "Share to Viva Engage", closeButtonAriaLabel: "Close", isFooterAtBottom: true, onRenderFooterContent: function () { return (React.createElement(react_2.Stack, { horizontal: true, tokens: { childrenGap: 8 } },
            React.createElement(react_2.PrimaryButton, { text: isPosting ? 'Posting…' : 'Post to Viva Engage', onClick: function () { void handlePost(); }, disabled: isPosting || !selectedGroupId || !!successMessage, styles: {
                    root: { backgroundColor: GREEN, borderColor: GREEN },
                    rootHovered: { backgroundColor: '#0b5e0b', borderColor: '#0b5e0b' },
                    rootDisabled: { backgroundColor: '#c8c6c4', borderColor: '#c8c6c4' }
                }, iconProps: { iconName: 'Share' } }),
            React.createElement(react_2.DefaultButton, { text: "Cancel", onClick: handleDismiss, disabled: isPosting }))); } },
        React.createElement(react_2.Stack, { tokens: { childrenGap: 16 }, style: { paddingTop: 8 } },
            React.createElement(react_2.Stack, { style: {
                    padding: '10px 14px',
                    borderRadius: 6,
                    backgroundColor: '#f3f9f3',
                    border: "1px solid #c3e6cb"
                } },
                React.createElement(react_2.Text, { variant: "tiny", style: { color: '#605e5c', marginBottom: 4 } }, "Sharing article"),
                React.createElement(react_2.Text, { variant: "medium", style: { fontWeight: 600, color: GREEN } }, articleTitle),
                React.createElement(react_2.Text, { variant: "tiny", style: { color: '#605e5c', wordBreak: 'break-all', marginTop: 4, opacity: 0.8 } }, articleUrl)),
            isLoadingGroups ? (React.createElement(react_2.Stack, { horizontal: true, tokens: { childrenGap: 8 }, verticalAlign: "center" },
                React.createElement(react_2.Spinner, { size: react_2.SpinnerSize.small }),
                React.createElement(react_2.Text, { variant: "small" }, "Loading your Viva Engage communities\u2026"))) : groupLoadError ? (React.createElement(react_2.Stack, { tokens: { childrenGap: 6 } },
                React.createElement(react_2.MessageBar, { messageBarType: react_2.MessageBarType.error, actions: React.createElement(react_2.DefaultButton, { text: "Retry", iconProps: { iconName: 'Refresh' }, onClick: onRetryLoadGroups, styles: { root: { minWidth: 70, height: 28, fontSize: 12 } } }) }, groupLoadError))) : (React.createElement(react_2.Dropdown, { label: "Community", placeholder: groups.length === 0 ? 'No Viva Engage communities found' : 'Select a Viva Engage community', options: groupOptions, selectedKey: selectedGroupId || undefined, onChange: function (_e, option) {
                    if (option)
                        setSelectedGroupId(String(option.key));
                }, disabled: groups.length === 0, required: true })),
            React.createElement(react_2.Stack, { tokens: { childrenGap: 4 } },
                React.createElement(react_2.Stack, { horizontal: true, horizontalAlign: "space-between", verticalAlign: "center" },
                    React.createElement(react_2.Label, { required: true }, "Description"),
                    (articleSummary === null || articleSummary === void 0 ? void 0 : articleSummary.trim()) && (React.createElement(react_2.Link, { onClick: function () { return setDescription(summaryToHtml(articleSummary.trim())); }, style: { fontSize: 12, color: GREEN } }, "Reset to article summary"))),
                React.createElement(react_quill_1.default, { theme: "snow", value: description, onChange: function (val) { return setDescription(val); }, modules: QUILL_MODULES, formats: QUILL_FORMATS, readOnly: isPosting, placeholder: "Edit the description that will appear in your Viva Engage post\u2026" }),
                React.createElement(react_2.Text, { variant: "tiny", style: { color: '#a19f9d' } }, "The article URL and \u201CShared via AI Curator\u201D attribution will be appended automatically.")),
            React.createElement(react_2.Stack, { style: {
                    padding: '10px 12px',
                    borderRadius: 6,
                    backgroundColor: '#f9f9f9',
                    border: '1px solid #edebe9'
                } },
                React.createElement(react_2.Stack, { horizontal: true, horizontalAlign: "space-between", verticalAlign: "center", style: { marginBottom: 6 } },
                    React.createElement(react_2.Text, { variant: "tiny", style: { color: '#605e5c', fontWeight: 600 } }, "Post preview"),
                    React.createElement(react_2.Link, { onClick: handleCopyToClipboard, style: { fontSize: 12, color: copied ? '#107C10' : '#605e5c' } }, copied ? '✓ Copied!' : 'Copy to clipboard')),
                React.createElement("div", { dangerouslySetInnerHTML: { __html: previewHtml }, style: { fontSize: 14, color: '#323130', lineHeight: '1.6', wordBreak: 'break-word' } })),
            errorMessage && (React.createElement(react_2.MessageBar, { messageBarType: react_2.MessageBarType.error, isMultiline: true, onDismiss: function () { return setErrorMessage(''); } }, errorMessage)),
            successMessage && (React.createElement(react_2.MessageBar, { messageBarType: react_2.MessageBarType.success, isMultiline: true }, successMessage)))));
};
exports.default = SharePanel;
//# sourceMappingURL=SharePanel.js.map