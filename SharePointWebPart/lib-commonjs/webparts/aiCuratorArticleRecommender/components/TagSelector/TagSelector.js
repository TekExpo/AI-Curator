"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var tslib_1 = require("tslib");
var React = tslib_1.__importStar(require("react"));
var react_1 = require("react");
var react_2 = require("@fluentui/react");
var TopicsService_1 = require("../../services/TopicsService");
var GREEN = '#107C10';
var CHIP_BORDER = '#c3e6cb';
var tagChipBase = {
    display: 'inline-block',
    padding: '4px 12px',
    margin: '4px',
    borderRadius: '16px',
    fontSize: '13px',
    fontWeight: 500,
    cursor: 'pointer',
    border: '1.5px solid',
    transition: 'all 0.15s ease',
    userSelect: 'none',
    lineHeight: '1.6'
};
/**
 * Tab 1 – My Interests
 * Renders a search box to discover topics via the suggest-topics API,
 * then renders selectable tag chips and a "Save My Interests" button.
 */
var TagSelector = function (props) {
    var savedTags = props.savedTags, successMessage = props.successMessage, onSave = props.onSave, isSaving = props.isSaving;
    // Parse savedTags into an array for auto-selection matching
    var savedTagsArray = (0, react_1.useMemo)(function () {
        return savedTags
            .split(',')
            .map(function (t) { return t.trim(); })
            .filter(function (t) { return t.length > 0; });
    }, [savedTags]);
    var _a = (0, react_1.useState)(''), searchQuery = _a[0], setSearchQuery = _a[1];
    var _b = (0, react_1.useState)([]), searchResults = _b[0], setSearchResults = _b[1];
    var _c = (0, react_1.useState)([]), selectedTags = _c[0], setSelectedTags = _c[1];
    var _d = (0, react_1.useState)(false), isSearching = _d[0], setIsSearching = _d[1];
    var _e = (0, react_1.useState)(false), searchError = _e[0], setSearchError = _e[1];
    var _f = (0, react_1.useState)(false), hasSearched = _f[0], setHasSearched = _f[1];
    var topicsService = (0, react_1.useRef)(new TopicsService_1.TopicsService());
    var initializedRef = (0, react_1.useRef)(false);
    // Seed selectedTags from saved data once it first arrives (e.g. after SP load)
    (0, react_1.useEffect)(function () {
        if (!initializedRef.current && savedTagsArray.length > 0) {
            initializedRef.current = true;
            setSelectedTags(savedTagsArray);
        }
    }, [savedTagsArray]);
    var handleSearch = function () { return tslib_1.__awaiter(void 0, void 0, void 0, function () {
        var topics_1, _a;
        return tslib_1.__generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    if (!searchQuery.trim())
                        return [2 /*return*/];
                    setIsSearching(true);
                    setSearchError(false);
                    _b.label = 1;
                case 1:
                    _b.trys.push([1, 3, 4, 5]);
                    return [4 /*yield*/, topicsService.current.getSuggestedTopics(searchQuery.trim())];
                case 2:
                    topics_1 = _b.sent();
                    setSearchResults(topics_1);
                    setHasSearched(true);
                    // Auto-select any returned topic that matches a previously saved tag
                    setSelectedTags(function (prev) {
                        var combined = tslib_1.__spreadArray([], prev, true);
                        topics_1.forEach(function (topic) {
                            if (savedTagsArray.some(function (s) { return s.toLowerCase() === topic.toLowerCase(); }) &&
                                !combined.some(function (t) { return t.toLowerCase() === topic.toLowerCase(); })) {
                                combined.push(topic);
                            }
                        });
                        return combined;
                    });
                    return [3 /*break*/, 5];
                case 3:
                    _a = _b.sent();
                    setSearchError(true);
                    setSearchResults([]);
                    setHasSearched(true);
                    return [3 /*break*/, 5];
                case 4:
                    setIsSearching(false);
                    return [7 /*endfinally*/];
                case 5: return [2 /*return*/];
            }
        });
    }); };
    var handleChipToggle = function (tag) {
        setSelectedTags(function (prev) {
            return prev.indexOf(tag) !== -1 ? prev.filter(function (t) { return t !== tag; }) : tslib_1.__spreadArray(tslib_1.__spreadArray([], prev, true), [tag], false);
        });
    };
    return (React.createElement(react_2.Stack, { tokens: { childrenGap: 16 }, style: { padding: '8px 0' } },
        React.createElement(react_2.Text, { variant: "mediumPlus", style: { fontWeight: 600, color: GREEN } }, "Select Your Interests"),
        React.createElement(react_2.Text, { variant: "small", style: { color: '#605e5c' } }, "Choose the topics you are interested in. We will recommend articles matching your selection."),
        savedTagsArray.length > 0 && (React.createElement(react_2.Stack, { tokens: { childrenGap: 8 }, style: {
                padding: '12px 14px',
                borderRadius: 8,
                backgroundColor: '#f3f9f3',
                border: "1px solid ".concat(CHIP_BORDER)
            } },
            React.createElement(react_2.Text, { variant: "small", style: { fontWeight: 600, color: GREEN } },
                "My Current Interests (",
                savedTagsArray.length,
                ")"),
            React.createElement(react_2.Text, { variant: "tiny", style: { color: '#605e5c' } }, "Click \u00D7 to remove a topic. Strikethrough means it will be removed when you save."),
            React.createElement("div", { style: { display: 'flex', flexWrap: 'wrap', gap: '4px', marginTop: 4 } }, savedTagsArray.map(function (tag) {
                var isActive = selectedTags.indexOf(tag) !== -1;
                return (React.createElement("span", { key: "saved-".concat(tag), style: {
                        display: 'inline-flex',
                        alignItems: 'center',
                        gap: 5,
                        padding: '3px 8px 3px 12px',
                        borderRadius: 16,
                        fontSize: 13,
                        fontWeight: 500,
                        border: '1.5px solid',
                        backgroundColor: isActive ? '#e8f5e9' : '#f3f2f1',
                        borderColor: isActive ? GREEN : '#c8c6c4',
                        color: isActive ? GREEN : '#a19f9d',
                        textDecoration: isActive ? 'none' : 'line-through',
                        opacity: isActive ? 1 : 0.75,
                        transition: 'all 0.15s ease'
                    } },
                    tag,
                    React.createElement("button", { onClick: function () {
                            return setSelectedTags(function (prev) {
                                return isActive
                                    ? prev.filter(function (t) { return t !== tag; })
                                    : tslib_1.__spreadArray(tslib_1.__spreadArray([], prev, true), [tag], false);
                            });
                        }, title: isActive ? "Remove \"".concat(tag, "\" from interests") : "Restore \"".concat(tag, "\""), style: {
                            background: 'none',
                            border: 'none',
                            cursor: 'pointer',
                            padding: '0 2px',
                            lineHeight: 1,
                            fontSize: 16,
                            fontWeight: 700,
                            color: isActive ? '#a4262c' : GREEN,
                            display: 'flex',
                            alignItems: 'center'
                        } }, isActive ? '\u00d7' : '+')));
            })))),
        React.createElement(react_2.Separator, null),
        React.createElement(react_2.Stack, { horizontal: true, tokens: { childrenGap: 8 }, verticalAlign: "end" },
            React.createElement(react_2.Stack.Item, { grow: true },
                React.createElement(react_2.TextField, { placeholder: "Search for topics...", value: searchQuery, onChange: function (_, val) { return setSearchQuery(val !== null && val !== void 0 ? val : ''); }, onKeyDown: function (e) {
                        if (e.key === 'Enter') {
                            void handleSearch();
                        }
                    } })),
            React.createElement(react_2.PrimaryButton, { text: "Search", onClick: function () { void handleSearch(); }, disabled: isSearching || !searchQuery.trim(), styles: {
                    root: { backgroundColor: GREEN, borderColor: GREEN, minWidth: 80 },
                    rootHovered: { backgroundColor: '#0b5e0b', borderColor: '#0b5e0b' },
                    rootDisabled: { backgroundColor: '#c8c6c4', borderColor: '#c8c6c4' }
                } })),
        isSearching && (React.createElement(react_2.Stack, { horizontalAlign: "center", style: { padding: '20px 0' } },
            React.createElement(react_2.Spinner, { size: react_2.SpinnerSize.medium, label: "Searching topics\u2026", labelPosition: "right" }))),
        !isSearching && searchError && (React.createElement(react_2.MessageBar, { messageBarType: react_2.MessageBarType.error, isMultiline: true }, "Unable to fetch topics. Please try again later.")),
        !isSearching && hasSearched && !searchError && searchResults.length === 0 && (React.createElement(react_2.MessageBar, { messageBarType: react_2.MessageBarType.info }, "No topics found for your search. Try different keywords.")),
        !isSearching && searchResults.length > 0 && (React.createElement(react_2.Stack, null,
            React.createElement("div", { style: { display: 'flex', flexWrap: 'wrap', gap: '2px' } }, searchResults.map(function (tag) {
                var isSelected = selectedTags.indexOf(tag) !== -1;
                return (React.createElement("span", { key: tag, role: "checkbox", "aria-checked": isSelected, tabIndex: 0, onClick: function () { return handleChipToggle(tag); }, onKeyDown: function (e) {
                        if (e.key === 'Enter' || e.key === ' ') {
                            e.preventDefault();
                            handleChipToggle(tag);
                        }
                    }, style: tslib_1.__assign(tslib_1.__assign({}, tagChipBase), { backgroundColor: '#ffffff', borderColor: isSelected ? GREEN : CHIP_BORDER, color: isSelected ? GREEN : '#323130', boxShadow: isSelected
                            ? '0 1px 4px rgba(16,124,16,0.3)'
                            : '0 1px 2px rgba(0,0,0,0.08)', outline: 'none' }), title: isSelected ? "Deselect \"".concat(tag, "\"") : "Select \"".concat(tag, "\"") }, tag));
            })),
            selectedTags.filter(function (t) { return searchResults.indexOf(t) !== -1; }).length > 0 && (React.createElement(react_2.Text, { variant: "small", style: { marginTop: 8, color: GREEN, fontWeight: 500 } },
                selectedTags.filter(function (t) { return searchResults.indexOf(t) !== -1; }).length,
                " topic",
                selectedTags.filter(function (t) { return searchResults.indexOf(t) !== -1; }).length !== 1 ? 's' : '',
                ' ',
                "selected")))),
        successMessage && (React.createElement(react_2.MessageBar, { messageBarType: react_2.MessageBarType.success, isMultiline: true }, successMessage)),
        React.createElement(react_2.PrimaryButton, { text: isSaving ? 'Saving…' : 'Save My Interests', iconProps: { iconName: 'Save' }, onClick: function () { return onSave(selectedTags); }, disabled: isSaving || selectedTags.length === 0, styles: {
                root: { backgroundColor: GREEN, borderColor: GREEN },
                rootHovered: { backgroundColor: '#0b5e0b', borderColor: '#0b5e0b' },
                rootDisabled: { backgroundColor: '#c8c6c4', borderColor: '#c8c6c4' }
            } }),
        selectedTags.length > 0 && (React.createElement(react_2.Stack, { style: {
                padding: '8px 12px',
                borderRadius: 6,
                backgroundColor: '#e8f5e9',
                border: "1px solid ".concat(CHIP_BORDER)
            } },
            React.createElement(react_2.Text, { variant: "small", style: { color: '#3a3a3a' } },
                React.createElement("strong", null, "Selected interests:"),
                " ",
                selectedTags.join(', '))))));
};
exports.default = TagSelector;
//# sourceMappingURL=TagSelector.js.map