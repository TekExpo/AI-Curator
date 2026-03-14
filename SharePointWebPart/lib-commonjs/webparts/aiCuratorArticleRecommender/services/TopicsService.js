"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.TopicsService = void 0;
var tslib_1 = require("tslib");
var BASE_URL = 'https://ai-curator.azurewebsites.net';
/**
 * Service for the external AI Curator API endpoints.
 * Uses native fetch() — no SPHttpClient or PnPjs dependency.
 */
var TopicsService = /** @class */ (function () {
    function TopicsService() {
    }
    /**
     * Fetches suggested topics from the suggest-topics endpoint.
     * @param query The user's search text
     * @returns Array of topic strings, or [] if none / missing from response
     */
    TopicsService.prototype.getSuggestedTopics = function (query) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var url, response, networkErr_1, data;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        url = "".concat(BASE_URL, "/suggest-topics?query=").concat(encodeURIComponent(query));
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, fetch(url)];
                    case 2:
                        response = _a.sent();
                        return [3 /*break*/, 4];
                    case 3:
                        networkErr_1 = _a.sent();
                        throw new Error("Network error fetching topics: ".concat(String(networkErr_1)));
                    case 4:
                        if (!response.ok) {
                            throw new Error("Topics API returned HTTP ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 5:
                        data = _a.sent();
                        if (!Array.isArray(data.topics))
                            return [2 /*return*/, []];
                        return [2 /*return*/, data.topics.filter(function (t) { return typeof t === 'string' && t.length > 0; })];
                }
            });
        });
    };
    /**
     * Fetches article recommendations from the articles endpoint.
     * @param selectedTags Comma-separated topic string from userPersonalization.SelectedTags
     * @param limit Maximum number of articles to return (default: 20)
     * @returns Array of IArticle objects
     */
    TopicsService.prototype.getArticles = function (selectedTags_1) {
        return tslib_1.__awaiter(this, arguments, void 0, function (selectedTags, limit) {
            var url, response, networkErr_2, data;
            if (limit === void 0) { limit = 20; }
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        url = "".concat(BASE_URL, "/articles") +
                            "?topics=".concat(encodeURIComponent(selectedTags), "&limit=").concat(limit);
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, fetch(url)];
                    case 2:
                        response = _a.sent();
                        return [3 /*break*/, 4];
                    case 3:
                        networkErr_2 = _a.sent();
                        throw new Error("Network error fetching articles: ".concat(String(networkErr_2)));
                    case 4:
                        if (!response.ok) {
                            throw new Error("Articles API returned HTTP ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 5:
                        data = _a.sent();
                        if (!Array.isArray(data.articles))
                            return [2 /*return*/, []];
                        return [2 /*return*/, data.articles.filter(function (a) {
                                return typeof (a === null || a === void 0 ? void 0 : a.title) === 'string' &&
                                    typeof (a === null || a === void 0 ? void 0 : a.url) === 'string';
                            })];
                }
            });
        });
    };
    return TopicsService;
}());
exports.TopicsService = TopicsService;
//# sourceMappingURL=TopicsService.js.map