declare interface IAiCuratorArticleRecommenderWebPartStrings {
  PropertyPaneDescription: string;
  DataSourceGroupName: string;
  PersonalizationGroupName: string;

  ListNameFieldLabel: string;

  UserPersonalizationListNameFieldLabel: string;
  ArticlesLimitFieldLabel: string;

  SearchPlaceholder: string;
  SearchButton: string;
  NoTopicsFound: string;
  TopicsFetchError: string;
  ArticlesFetchError: string;
  NoArticlesFound: string;
  NoInterestsSaved: string;
  EmptyInterests: string;
}

declare module 'AiCuratorArticleRecommenderWebPartStrings' {
  const strings: IAiCuratorArticleRecommenderWebPartStrings;
  export = strings;
}
