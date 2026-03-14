"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.SharePointService = void 0;
var tslib_1 = require("tslib");
var sp_http_1 = require("@microsoft/sp-http");
/**
 * Service for all SharePoint REST API operations.
 * Uses SPHttpClient directly — no PnPjs dependency required.
 */
var SharePointService = /** @class */ (function () {
    function SharePointService(context) {
        this._spHttpClient = context.spHttpClient;
        this._webUrl = context.pageContext.web.absoluteUrl;
    }
    /**
     * Returns numeric Id and LoginName of the currently logged-in user.
     */
    SharePointService.prototype.getCurrentUser = function () {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var url, response, data;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        url = "".concat(this._webUrl, "/_api/web/currentUser?$select=Id,LoginName,Title");
                        return [4 /*yield*/, this._spHttpClient.get(url, sp_http_1.SPHttpClient.configurations.v1)];
                    case 1:
                        response = _a.sent();
                        if (!response.ok) {
                            throw new Error("Failed to get current user: HTTP ".concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _a.sent();
                        return [2 /*return*/, { Id: data.Id, LoginName: data.LoginName, Title: data.Title }];
                }
            });
        });
    };
    /**
     * Queries userPersonalization by numeric UserId.
     * Returns null if no record exists for this user.
     */
    SharePointService.prototype.getUserPersonalizationByUserId = function (listName, userId) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var url, response, data, items;
            var _a, _b, _c;
            return tslib_1.__generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        url = "".concat(this._webUrl, "/_api/web/lists/getbytitle('").concat(encodeURIComponent(listName), "')/items") +
                            "?$select=Id,SelectedTags,SavedLinks&$filter=UserId eq ".concat(userId, "&$top=1");
                        return [4 /*yield*/, this._spHttpClient.get(url, sp_http_1.SPHttpClient.configurations.v1)];
                    case 1:
                        response = _d.sent();
                        if (!response.ok) {
                            throw new Error("Failed to query \"".concat(listName, "\" for userId ").concat(userId, ": HTTP ").concat(response.status));
                        }
                        return [4 /*yield*/, response.json()];
                    case 2:
                        data = _d.sent();
                        items = (_a = data.value) !== null && _a !== void 0 ? _a : [];
                        if (items.length === 0)
                            return [2 /*return*/, null];
                        return [2 /*return*/, {
                                itemId: items[0].Id,
                                SelectedTags: (_b = items[0].SelectedTags) !== null && _b !== void 0 ? _b : '',
                                SavedLinks: (_c = items[0].SavedLinks) !== null && _c !== void 0 ? _c : ''
                            }];
                }
            });
        });
    };
    /**
     * Creates a new userPersonalization record for the current user.
     */
    SharePointService.prototype.createUserPersonalization = function (listName, loginName, userId, selectedTags) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var url, body, response, errorText, created;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        url = "".concat(this._webUrl, "/_api/web/lists/getbytitle('").concat(encodeURIComponent(listName), "')/items");
                        body = JSON.stringify({
                            Title: loginName,
                            UserId: userId,
                            SelectedTags: selectedTags,
                            SavedLinks: ''
                        });
                        return [4 /*yield*/, this._spHttpClient.post(url, sp_http_1.SPHttpClient.configurations.v1, {
                                headers: {
                                    Accept: 'application/json;odata=nometadata',
                                    'Content-Type': 'application/json;odata=nometadata',
                                    'odata-version': ''
                                },
                                body: body
                            })];
                    case 1:
                        response = _a.sent();
                        if (!!response.ok) return [3 /*break*/, 3];
                        return [4 /*yield*/, response.text().catch(function () { return ''; })];
                    case 2:
                        errorText = _a.sent();
                        throw new Error("Failed to create record in \"".concat(listName, "\": HTTP ").concat(response.status, " \u2014 ").concat(errorText.substring(0, 300)));
                    case 3: return [4 /*yield*/, response.json()];
                    case 4:
                        created = _a.sent();
                        return [2 /*return*/, { itemId: created.Id }];
                }
            });
        });
    };
    /**
     * Updates SelectedTags for an existing userPersonalization record by item ID.
     */
    SharePointService.prototype.updateSelectedTags = function (listName, itemId, selectedTags) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var url, body, response;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        url = "".concat(this._webUrl, "/_api/web/lists/getbytitle('").concat(encodeURIComponent(listName), "')/items(").concat(itemId, ")");
                        body = JSON.stringify({ SelectedTags: selectedTags });
                        return [4 /*yield*/, this._spHttpClient.fetch(url, sp_http_1.SPHttpClient.configurations.v1, {
                                method: 'PATCH',
                                headers: {
                                    Accept: 'application/json;odata=nometadata',
                                    'Content-Type': 'application/json;odata=nometadata',
                                    'odata-version': '',
                                    'IF-MATCH': '*',
                                    'X-HTTP-Method': 'MERGE'
                                },
                                body: body
                            })];
                    case 1:
                        response = _a.sent();
                        if (!response.ok && response.status !== 204) {
                            throw new Error("Failed to update SelectedTags in \"".concat(listName, "\" item ").concat(itemId, ": HTTP ").concat(response.status));
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Appends an article URL to SavedLinks (comma-separated, no duplicates).
     */
    SharePointService.prototype.saveArticleLink = function (listName, itemId, currentSavedLinks, newArticleUrl) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var existing, updatedLinks, url, body, response;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        existing = currentSavedLinks
                            .split(',')
                            .map(function (l) { return l.trim(); })
                            .filter(function (l) { return l.length > 0; });
                        if (existing.indexOf(newArticleUrl.trim()) !== -1) {
                            return [2 /*return*/]; // already saved — no-op
                        }
                        existing.push(newArticleUrl.trim());
                        updatedLinks = existing.join(',');
                        url = "".concat(this._webUrl, "/_api/web/lists/getbytitle('").concat(encodeURIComponent(listName), "')/items(").concat(itemId, ")");
                        body = JSON.stringify({ SavedLinks: updatedLinks });
                        return [4 /*yield*/, this._spHttpClient.fetch(url, sp_http_1.SPHttpClient.configurations.v1, {
                                method: 'PATCH',
                                headers: {
                                    Accept: 'application/json;odata=nometadata',
                                    'Content-Type': 'application/json;odata=nometadata',
                                    'odata-version': '',
                                    'IF-MATCH': '*',
                                    'X-HTTP-Method': 'MERGE'
                                },
                                body: body
                            })];
                    case 1:
                        response = _a.sent();
                        if (!response.ok && response.status !== 204) {
                            throw new Error("Failed to save link in \"".concat(listName, "\" item ").concat(itemId, ": HTTP ").concat(response.status));
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    /**
     * Removes an article URL from SavedLinks (comma-separated).
     */
    SharePointService.prototype.removeSavedLink = function (listName, itemId, currentSavedLinks, urlToRemove) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var updated, url, body, response;
            return tslib_1.__generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        updated = currentSavedLinks
                            .split(',')
                            .map(function (l) { return l.trim(); })
                            .filter(function (l) { return l.length > 0 && l !== urlToRemove.trim(); })
                            .join(',');
                        url = "".concat(this._webUrl, "/_api/web/lists/getbytitle('").concat(encodeURIComponent(listName), "')/items(").concat(itemId, ")");
                        body = JSON.stringify({ SavedLinks: updated });
                        return [4 /*yield*/, this._spHttpClient.fetch(url, sp_http_1.SPHttpClient.configurations.v1, {
                                method: 'PATCH',
                                headers: {
                                    Accept: 'application/json;odata=nometadata',
                                    'Content-Type': 'application/json;odata=nometadata',
                                    'odata-version': '',
                                    'IF-MATCH': '*',
                                    'X-HTTP-Method': 'MERGE'
                                },
                                body: body
                            })];
                    case 1:
                        response = _a.sent();
                        if (!response.ok && response.status !== 204) {
                            throw new Error("Failed to remove link from \"".concat(listName, "\" item ").concat(itemId, ": HTTP ").concat(response.status));
                        }
                        return [2 /*return*/];
                }
            });
        });
    };
    return SharePointService;
}());
exports.SharePointService = SharePointService;
//# sourceMappingURL=SharePointService.js.map