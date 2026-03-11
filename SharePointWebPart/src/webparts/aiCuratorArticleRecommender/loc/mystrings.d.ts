declare interface IAiCuratorArticleRecommenderWebPartStrings {
  PropertyPaneDescription: string;
  LlmConfigGroupName: string;
  DataSourceGroupName: string;
  PersonalizationGroupName: string;

  OpenAiModelFieldLabel: string;
  OpenAiSystemPromptFieldLabel: string;
  MaxArticlesFieldLabel: string;
  EnableCachingFieldLabel: string;

  ListNameFieldLabel: string;

  UserPersonalizationListNameFieldLabel: string;
  VivaEngageEnabledFieldLabel: string;
  YammerClientIdFieldLabel: string;
}

declare module 'AiCuratorArticleRecommenderWebPartStrings' {
  const strings: IAiCuratorArticleRecommenderWebPartStrings;
  export = strings;
}
