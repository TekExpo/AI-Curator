declare interface IAiCuratorArticleRecommenderWebPartStrings {
  PropertyPaneDescription: string;
  LlmConfigGroupName: string;
  DataSourceGroupName: string;

  OpenAiModelFieldLabel: string;
  OpenAiSystemPromptFieldLabel: string;
  MaxArticlesFieldLabel: string;
  EnableCachingFieldLabel: string;

  ListNameFieldLabel: string;
  KeywordColumnNameFieldLabel: string;
  SiteUrlFieldLabel: string;
}

declare module 'AiCuratorArticleRecommenderWebPartStrings' {
  const strings: IAiCuratorArticleRecommenderWebPartStrings;
  export = strings;
}
