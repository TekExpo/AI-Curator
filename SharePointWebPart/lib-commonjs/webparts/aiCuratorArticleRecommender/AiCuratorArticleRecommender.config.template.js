"use strict";
// Copy this file to AiCuratorArticleRecommender.config.ts and fill in the real values.
// AiCuratorArticleRecommender.config.ts is gitignored and must NEVER be committed.
//
// ⚠️  SECURITY NOTE: This web part runs in the browser. Any value placed here
// will be visible in the compiled JavaScript bundle. For production use, replace
// the direct OpenAI call with a server-side proxy (e.g. an Azure Function) so
// the API key never leaves the server.
Object.defineProperty(exports, "__esModule", { value: true });
exports.aiCuratorArticleRecommenderConfig = void 0;
exports.aiCuratorArticleRecommenderConfig = {
    llmEndpointUrl: 'https://<resource>.cognitiveservices.azure.com/openai/deployments/<model>/chat/completions?api-version=2025-01-01-preview',
    openAiApiKey: '<your-azure-openai-api-key>'
};
//# sourceMappingURL=AiCuratorArticleRecommender.config.template.js.map