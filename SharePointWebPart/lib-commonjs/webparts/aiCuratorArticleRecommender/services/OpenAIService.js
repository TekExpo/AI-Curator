"use strict";
// NOTE: OpenAIService is no longer called by AiCuratorArticleRecommender Tab 2.
// Article recommendations are now fetched from the external articles endpoint via TopicsService.
// This file is retained for reference only.
Object.defineProperty(exports, "__esModule", { value: true });
exports.OpenAIService = void 0;
var tslib_1 = require("tslib");
// ------------------------------------------------------------------
// Internal helpers (moved from AiCuratorArticleRecommender.tsx)
// ------------------------------------------------------------------
var CACHE_KEY_PREFIX = 'ai-curator-cache-';
var CACHE_TTL_MS = 15 * 60 * 1000; // 15 minutes
var DEFAULT_OPENAI_SYSTEM_PROMPT = [
    'You recommend articles based on a list of keywords.',
    'Return only valid JSON.',
    'The JSON must be an array of objects with this exact schema:',
    '[{"title":"string","url":"https://example.com","summary":"string"}]',
    'Use real-looking article titles and absolute URLs.',
    'Return at most {{maxArticles}} items.'
].join(' ');
function hashString(input) {
    var hash = 5381;
    for (var i = 0; i < input.length; i++) {
        // eslint-disable-next-line no-bitwise
        hash = ((hash << 5) + hash) + input.charCodeAt(i);
    }
    return hash.toString(36);
}
function isValidArticleArray(data) {
    if (!Array.isArray(data))
        return false;
    return data.every(function (item) {
        return typeof item === 'object' &&
            item !== null &&
            typeof item.title === 'string' &&
            typeof item.url === 'string' &&
            (typeof item.summary === 'string' ||
                typeof item.description === 'string');
    });
}
function extractJsonPayload(text) {
    var _a, _b;
    var fencedMatch = text.match(/```(?:json)?\s*([\s\S]*?)```/i);
    var candidate = (_b = (_a = fencedMatch === null || fencedMatch === void 0 ? void 0 : fencedMatch[1]) === null || _a === void 0 ? void 0 : _a.trim()) !== null && _b !== void 0 ? _b : text.trim();
    var arrayStart = candidate.indexOf('[');
    var arrayEnd = candidate.lastIndexOf(']');
    if (arrayStart >= 0 && arrayEnd > arrayStart) {
        return candidate.substring(arrayStart, arrayEnd + 1);
    }
    return candidate;
}
function normalizeArticleArray(raw) {
    return raw.map(function (item) {
        var _a, _b, _c, _d;
        var r = item;
        return {
            title: String((_a = r.title) !== null && _a !== void 0 ? _a : ''),
            url: String((_b = r.url) !== null && _b !== void 0 ? _b : ''),
            description: String((_d = (_c = r.summary) !== null && _c !== void 0 ? _c : r.description) !== null && _d !== void 0 ? _d : '')
        };
    });
}
function parseArticleArrayFromText(text) {
    try {
        var parsed = JSON.parse(extractJsonPayload(text));
        if (isValidArticleArray(parsed))
            return normalizeArticleArray(parsed);
        if (typeof parsed === 'object' && parsed !== null && 'articles' in parsed) {
            var articles = parsed.articles;
            if (isValidArticleArray(articles))
                return normalizeArticleArray(articles);
        }
    }
    catch (_a) {
        // ignore
    }
    return null;
}
function parseOpenAiResponse(data) {
    var _a, _b, _c, _d, _e, _f;
    if (isValidArticleArray(data))
        return normalizeArticleArray(data);
    if (typeof data === 'object' && data !== null && 'articles' in data) {
        var articles = data.articles;
        if (isValidArticleArray(articles))
            return normalizeArticleArray(articles);
    }
    var chatContent = (_c = (_b = (_a = data === null || data === void 0 ? void 0 : data.choices) === null || _a === void 0 ? void 0 : _a[0]) === null || _b === void 0 ? void 0 : _b.message) === null || _c === void 0 ? void 0 : _c.content;
    if (typeof chatContent === 'string')
        return parseArticleArrayFromText(chatContent);
    if (Array.isArray(chatContent)) {
        var contentText = chatContent
            .map(function (part) {
            if (typeof part === 'string')
                return part;
            if (typeof part === 'object' && part !== null && 'text' in part) {
                var t = part.text;
                return typeof t === 'string' ? t : '';
            }
            return '';
        })
            .join('');
        if (contentText)
            return parseArticleArrayFromText(contentText);
    }
    var responsesText = (_f = (_e = (_d = data === null || data === void 0 ? void 0 : data.output) === null || _d === void 0 ? void 0 : _d[0]) === null || _e === void 0 ? void 0 : _e.content) === null || _f === void 0 ? void 0 : _f.map(function (part) { var _a; return (_a = part.text) !== null && _a !== void 0 ? _a : ''; }).join('');
    if (responsesText)
        return parseArticleArrayFromText(responsesText);
    return null;
}
// ------------------------------------------------------------------
// OpenAIService
// ------------------------------------------------------------------
/**
 * Handles all OpenAI API communication.
 * Keywords are always the raw SelectedTags string from userPersonalization.
 */
var OpenAIService = /** @class */ (function () {
    function OpenAIService(endpointUrl, apiKey) {
        this._endpointUrl = endpointUrl;
        this._apiKey = apiKey;
    }
    OpenAIService.prototype.getArticleRecommendations = function (tags, model, systemPrompt, maxArticles, enableCaching, signal) {
        return tslib_1.__awaiter(this, void 0, void 0, function () {
            var endpointUrl, resolvedPrompt, cacheKey, cached, parsed, isOpenAiEndpoint, userPrompt, payload, headers, response, errorBody, data, _a, parsedArticles, cacheKey;
            var _b, _c;
            return tslib_1.__generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        endpointUrl = (_b = this._endpointUrl) === null || _b === void 0 ? void 0 : _b.trim();
                        if (!endpointUrl) {
                            throw new Error('LLM Endpoint URL is not configured. Please set it in the web part config file.');
                        }
                        resolvedPrompt = ((systemPrompt === null || systemPrompt === void 0 ? void 0 : systemPrompt.trim().length) > 0 ? systemPrompt.trim() : DEFAULT_OPENAI_SYSTEM_PROMPT)
                            .replace(/\{\{keywords\}\}/g, tags)
                            .replace(/\{\{maxArticles\}\}/g, String(maxArticles));
                        if (enableCaching) {
                            cacheKey = CACHE_KEY_PREFIX +
                                hashString("".concat(endpointUrl, "|").concat(model, "|").concat(resolvedPrompt, "|").concat(tags, "|").concat(maxArticles));
                            try {
                                cached = sessionStorage.getItem(cacheKey);
                                if (cached) {
                                    parsed = JSON.parse(cached);
                                    if (Date.now() - parsed.timestamp < CACHE_TTL_MS) {
                                        return [2 /*return*/, parsed.data];
                                    }
                                    sessionStorage.removeItem(cacheKey);
                                }
                            }
                            catch (_e) {
                                // sessionStorage unavailable — continue
                            }
                        }
                        isOpenAiEndpoint = /api\.openai\.com|openai\.azure\.com/i.test(endpointUrl) ||
                            endpointUrl.includes('/chat/completions') ||
                            endpointUrl.includes('/responses');
                        userPrompt = [
                            "Keywords: ".concat(tags),
                            "Maximum articles: ".concat(maxArticles),
                            'Return article recommendations for these keywords with working URLs and concise summaries.'
                        ].join('\n');
                        payload = isOpenAiEndpoint
                            ? tslib_1.__assign({ messages: [
                                    { role: 'system', content: resolvedPrompt },
                                    { role: 'user', content: userPrompt }
                                ], temperature: 0.2 }, ((model === null || model === void 0 ? void 0 : model.trim().length) > 0 ? { model: model.trim() } : {})) : { keywords: tags, maxResults: maxArticles, prompt: resolvedPrompt };
                        headers = {
                            'Content-Type': 'application/json',
                            Accept: 'application/json'
                        };
                        if (((_c = this._apiKey) === null || _c === void 0 ? void 0 : _c.trim().length) > 0) {
                            if (/openai\.azure\.com/i.test(endpointUrl)) {
                                headers['api-key'] = this._apiKey.trim();
                            }
                            else {
                                headers.Authorization = "Bearer ".concat(this._apiKey.trim());
                            }
                        }
                        return [4 /*yield*/, fetch(endpointUrl, {
                                method: 'POST',
                                headers: headers,
                                body: JSON.stringify(payload),
                                signal: signal
                            })];
                    case 1:
                        response = _d.sent();
                        if (!!response.ok) return [3 /*break*/, 3];
                        return [4 /*yield*/, response.text().catch(function () { return 'No response body'; })];
                    case 2:
                        errorBody = _d.sent();
                        throw new Error("LLM endpoint returned HTTP ".concat(response.status, ": ").concat(response.statusText, ". ") +
                            "Body: ".concat(errorBody.substring(0, 200)));
                    case 3:
                        _d.trys.push([3, 5, , 6]);
                        return [4 /*yield*/, response.json()];
                    case 4:
                        data = _d.sent();
                        return [3 /*break*/, 6];
                    case 5:
                        _a = _d.sent();
                        throw new Error('LLM endpoint returned invalid JSON. Expected an array of ' +
                            '{ title, url, summary } objects.');
                    case 6:
                        parsedArticles = parseOpenAiResponse(data);
                        if (!parsedArticles) {
                            throw new Error('LLM response does not match expected schema. Expected array of ' +
                                '[{ "title": string, "url": string, "summary": string }, ...]. ' +
                                "Received: ".concat(JSON.stringify(data).substring(0, 300)));
                        }
                        if (enableCaching) {
                            cacheKey = CACHE_KEY_PREFIX +
                                hashString("".concat(endpointUrl, "|").concat(model, "|").concat(resolvedPrompt, "|").concat(tags, "|").concat(maxArticles));
                            try {
                                sessionStorage.setItem(cacheKey, JSON.stringify({ timestamp: Date.now(), data: parsedArticles }));
                            }
                            catch (_f) {
                                // sessionStorage full or unavailable
                            }
                        }
                        return [2 /*return*/, parsedArticles];
                }
            });
        });
    };
    return OpenAIService;
}());
exports.OpenAIService = OpenAIService;
//# sourceMappingURL=OpenAIService.js.map