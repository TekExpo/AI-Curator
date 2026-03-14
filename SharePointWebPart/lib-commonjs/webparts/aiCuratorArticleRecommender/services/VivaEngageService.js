"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.VivaEngageService = void 0;
var tslib_1 = require("tslib");
/**
 * Service for Viva Engage (Yammer) operations via Microsoft Graph API.
 * Uses MSGraphClientV3 for all calls.
 */
var VivaEngageService = /** @class */ (function () {
    function VivaEngageService(graphClient) {
        this._graphClient = graphClient;
    }
    /**
     * Returns Viva Engage communities the current user belongs to.
     * Filters M365 groups to those provisioned with a Yammer/Viva Engage backend.
     */
    VivaEngageService.prototype.getYammerGroups = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var result, groups, err_1, msg;
            var _a;
            return tslib_1.__generator(this, function (_b) {
                switch (_b.label) {
                    case 0:
                        _b.trys.push([0, 2, , 3]);
                        return [4 /*yield*/, this._graphClient
                                .api('/groups')
                                .filter("resourceProvisioningOptions/Any(x:x eq 'Yammer')")
                                .select('id,displayName')
                                .top(100)
                                .get()];
                    case 1:
                        result = _b.sent();
                        groups = (_a = result.value) !== null && _a !== void 0 ? _a : [];
                        return [2 /*return*/, groups.map(function (g) { return ({ id: g.id, name: g.displayName }); })];
                    case 2:
                        err_1 = _b.sent();
                        msg = err_1 instanceof Error ? err_1.message : String(err_1);
                        throw new Error("Failed to fetch Viva Engage communities: ".concat(msg));
                    case 3: return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Posts a message to a Viva Engage / Yammer group via Microsoft Graph.
     * Message body: userComments + articleUrl + attribution.
     */
    VivaEngageService.prototype.postToGroup = function (groupId, articleUrl, descriptionHtml) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var safeUrl, htmlContent, err_2, msg;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        safeUrl = articleUrl
                            .replace(/&/g, '&amp;')
                            .replace(/</g, '&lt;')
                            .replace(/>/g, '&gt;')
                            .replace(/"/g, '&quot;');
                        htmlContent = ((descriptionHtml === null || descriptionHtml === void 0 ? void 0 : descriptionHtml.trim()) || '') +
                            "<p><a href=\"".concat(safeUrl, "\">").concat(safeUrl, "</a></p>") +
                            '<p><em>Shared via AI Curator \u2013 Article Recommender</em></p>';
                        _a.label = 1;
                    case 1:
                        _a.trys.push([1, 3, , 4]);
                        return [4 /*yield*/, this._graphClient
                                .api("/groups/".concat(groupId, "/threads"))
                                .post({
                                topic: 'AI Curator \u2013 Shared Article',
                                posts: [
                                    {
                                        post: {
                                            body: {
                                                contentType: 'html',
                                                content: htmlContent
                                            }
                                        }
                                    }
                                ]
                            })];
                    case 2:
                        _a.sent();
                        return [3 /*break*/, 4];
                    case 3:
                        err_2 = _a.sent();
                        msg = err_2 instanceof Error ? err_2.message : String(err_2);
                        throw new Error("Failed to post to Viva Engage group: ".concat(msg));
                    case 4: return [2 /*return*/];
                }
            });
        });
    };
    return VivaEngageService;
}());
exports.VivaEngageService = VivaEngageService;
//# sourceMappingURL=VivaEngageService.js.map