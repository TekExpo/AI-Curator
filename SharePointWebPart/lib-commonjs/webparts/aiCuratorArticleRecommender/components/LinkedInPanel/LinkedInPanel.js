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
var LINKEDIN_BLUE = '#0A66C2';
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
 * Slide-in panel for sharing an article to LinkedIn.
 * The user edits a description, then clicks "Share on LinkedIn" which
 * opens LinkedIn's share composer pre-populated with the article title,
 * description, and URL — the user simply clicks "Post" inside LinkedIn.
 */
var LinkedInPanel = function (props) {
    var isOpen = props.isOpen, articleUrl = props.articleUrl, articleTitle = props.articleTitle, articleSummary = props.articleSummary, onDismiss = props.onDismiss;
    var _a = (0, react_1.useState)(''), description = _a[0], setDescription = _a[1];
    var _b = (0, react_1.useState)(false), isSharing = _b[0], setIsSharing = _b[1];
    // Reset and pre-populate whenever the panel opens
    (0, react_1.useEffect)(function () {
        if (isOpen) {
            setDescription((articleSummary === null || articleSummary === void 0 ? void 0 : articleSummary.trim()) ? summaryToHtml(articleSummary.trim()) : '');
            setIsSharing(false);
        }
    }, [isOpen]);
    var handleDismiss = function () {
        setDescription('');
        setIsSharing(false);
        onDismiss();
    };
    var handleShare = function () {
        setIsSharing(true);
        // Strip HTML to get plain text — LinkedIn's shareArticle endpoint accepts plain text
        // in its summary param and renders it in the composer.
        var plainText = htmlToPlainText(description);
        // shareArticle?mini=true pre-populates LinkedIn's post composer with the
        // title, summary, and URL. The user just clicks "Post" in the LinkedIn window.
        var params = new URLSearchParams({
            mini: 'true',
            url: articleUrl,
            title: articleTitle,
            summary: plainText
        });
        var linkedInUrl = "https://www.linkedin.com/shareArticle?".concat(params.toString());
        window.open(linkedInUrl, '_blank', 'noopener,noreferrer');
        // Reset sharing state after a short delay so button becomes re-usable
        setTimeout(function () { return setIsSharing(false); }, 1500);
    };
    var previewHtml = (description.trim() || '') +
        (articleUrl
            ? "<p><a href=\"".concat(escapeHtml(articleUrl), "\" target=\"_blank\" rel=\"noopener noreferrer\">").concat(escapeHtml(articleUrl), "</a></p>")
            : '') +
        '<p><em>Shared via AI Curator \u2013 Article Recommender</em></p>';
    return (React.createElement(react_2.Panel, { isOpen: isOpen, onDismiss: handleDismiss, type: react_2.PanelType.medium, headerText: "Share on LinkedIn", closeButtonAriaLabel: "Close", isFooterAtBottom: true, onRenderFooterContent: function () { return (React.createElement(react_2.Stack, { horizontal: true, tokens: { childrenGap: 8 } },
            React.createElement(react_2.PrimaryButton, { text: isSharing ? 'Opening LinkedIn…' : 'Share on LinkedIn', onClick: handleShare, disabled: isSharing, styles: {
                    root: { backgroundColor: LINKEDIN_BLUE, borderColor: LINKEDIN_BLUE },
                    rootHovered: { backgroundColor: '#004182', borderColor: '#004182' },
                    rootDisabled: { backgroundColor: '#c8c6c4', borderColor: '#c8c6c4' }
                }, onRenderIcon: function () { return (React.createElement("svg", { xmlns: "http://www.w3.org/2000/svg", viewBox: "0 0 24 24", width: "16", height: "16", "aria-hidden": "true", focusable: "false", style: { marginRight: 6, fill: 'currentColor' } },
                    React.createElement("path", { d: "M20.447 20.452h-3.554v-5.569c0-1.328-.027-3.037-1.852-3.037-1.853 0-2.136 1.445-2.136 2.939v5.667H9.351V9h3.414v1.561h.046c.477-.9 1.637-1.85 3.37-1.85 3.601 0 4.267 2.37 4.267 5.455v6.286zM5.337 7.433a2.062 2.062 0 0 1-2.063-2.065 2.064 2.064 0 1 1 2.063 2.065zm1.782 13.019H3.555V9h3.564v11.452zM22.225 0H1.771C.792 0 0 .774 0 1.729v20.542C0 23.227.792 24 1.771 24h20.451C23.2 24 24 23.227 24 22.271V1.729C24 .774 23.2 0 22.222 0h.003z" }))); } }),
            React.createElement(react_2.DefaultButton, { text: "Cancel", onClick: handleDismiss, disabled: isSharing }))); } },
        React.createElement(react_2.Stack, { tokens: { childrenGap: 16 }, style: { paddingTop: 8 } },
            React.createElement(react_2.Stack, { style: {
                    padding: '10px 14px',
                    borderRadius: 6,
                    backgroundColor: '#e8f0fb',
                    border: "1px solid #b3caf0"
                } },
                React.createElement(react_2.Text, { variant: "tiny", style: { color: '#605e5c', marginBottom: 4 } }, "Sharing article"),
                React.createElement(react_2.Text, { variant: "medium", style: { fontWeight: 600, color: LINKEDIN_BLUE } }, articleTitle),
                React.createElement(react_2.Text, { variant: "tiny", style: { color: '#605e5c', wordBreak: 'break-all', marginTop: 4, opacity: 0.8 } }, articleUrl)),
            React.createElement(react_2.Stack, { tokens: { childrenGap: 4 } },
                React.createElement(react_2.Stack, { horizontal: true, horizontalAlign: "space-between", verticalAlign: "center" },
                    React.createElement(react_2.Label, null, "Description"),
                    (articleSummary === null || articleSummary === void 0 ? void 0 : articleSummary.trim()) && (React.createElement(react_2.Link, { onClick: function () { return setDescription(summaryToHtml(articleSummary.trim())); }, style: { fontSize: 12, color: LINKEDIN_BLUE } }, "Reset to article summary"))),
                React.createElement(react_quill_1.default, { theme: "snow", value: description, onChange: function (val) { return setDescription(val); }, modules: QUILL_MODULES, formats: QUILL_FORMATS, readOnly: isSharing, placeholder: "Edit the description for your LinkedIn post\u2026" }),
                React.createElement(react_2.Text, { variant: "tiny", style: { color: '#a19f9d' } }, "Edit the description above. When you click \"Share on LinkedIn\", LinkedIn\u2019s post composer will open with your title, description, and article link already filled in \u2014 just click Post.")),
            React.createElement(react_2.Stack, { style: {
                    padding: '10px 12px',
                    borderRadius: 6,
                    backgroundColor: '#f9f9f9',
                    border: '1px solid #edebe9'
                } },
                React.createElement(react_2.Text, { variant: "tiny", style: { color: '#605e5c', fontWeight: 600, marginBottom: 6 } }, "Post preview"),
                React.createElement("div", { dangerouslySetInnerHTML: { __html: previewHtml }, style: { fontSize: 14, color: '#323130', lineHeight: '1.6', wordBreak: 'break-word' } })))));
};
exports.default = LinkedInPanel;
//# sourceMappingURL=LinkedInPanel.js.map