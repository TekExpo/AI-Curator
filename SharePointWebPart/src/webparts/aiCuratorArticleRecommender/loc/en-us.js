define([], function() {
  return {
    "PropertyPaneDescription": "Configure the AI Curator – Article Recommender web part. The LLM endpoint and API key are set in the config file. Keywords come from the userPersonalization list at runtime — never from this property pane.",
    "LlmConfigGroupName": "LLM Configuration",
    "DataSourceGroupName": "SharePoint Data Source",
    "PersonalizationGroupName": "Personalization & Sharing",

    "OpenAiModelFieldLabel": "OpenAI Model",
    "OpenAiSystemPromptFieldLabel": "OpenAI System Prompt",
    "MaxArticlesFieldLabel": "Max Number of Articles",
    "EnableCachingFieldLabel": "Enable Response Caching",

    "ListNameFieldLabel": "SharePoint List Name",

    "UserPersonalizationListNameFieldLabel": "User Personalization List Name",
    "VivaEngageEnabledFieldLabel": "Enable Viva Engage Sharing",
    "YammerClientIdFieldLabel": "Yammer App Client ID"
  }
});
