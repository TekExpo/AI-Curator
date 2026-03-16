import * as React from 'react';
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  ActionButton
} from '@fluentui/react';
import ArticleCard from './ArticleCard';
import { IArticleListProps } from './IArticleListProps';

/**
 * Tab 2 – Recommended Articles
 * Renders the AI-powered article recommendation list.
 */
const ArticleList: React.FC<IArticleListProps> = (props) => {
  const {
    articles,
    isLoading,
    errorMessage,
    infoMessage,
    savedArticleUrls,
    onSaveArticle,
    onShareArticle,
    onLinkedInShareArticle,
    onReload
  } = props;

  if (isLoading) {
    return (
      <Stack
        horizontalAlign="center"
        verticalAlign="center"
        style={{ minHeight: 200, padding: 40 }}
      >
        <Spinner
          size={SpinnerSize.large}
          label="Fetching article recommendations…"
          labelPosition="bottom"
        />
      </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 12 }} style={{ padding: '8px 0' }}>
      {/* Toolbar row */}
      <Stack horizontal horizontalAlign="end">
        <ActionButton
          iconProps={{ iconName: 'Refresh' }}
          onClick={onReload}
          disabled={isLoading}
          ariaLabel="Reload recommendations"
          styles={{ root: { color: '#605e5c' }, rootHovered: { color: '#107C10' } }}
        >
          Reload
        </ActionButton>
      </Stack>
      {errorMessage && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          {errorMessage}
        </MessageBar>
      )}

      {infoMessage && !errorMessage && (
        <MessageBar messageBarType={MessageBarType.info} isMultiline>
          {infoMessage}
        </MessageBar>
      )}

      {!isLoading && !errorMessage && !infoMessage && articles.length === 0 && (
        <MessageBar messageBarType={MessageBarType.info}>
          No article recommendations were returned. Try updating your interests in the My Interests tab.
        </MessageBar>
      )}

      {articles.map((article, index) => (
        <ArticleCard
          key={`article-${index}-${article.url}`}
          article={article}
          index={index}
          isSaved={savedArticleUrls.indexOf(article.url) !== -1}
          onSave={onSaveArticle}
          onShare={onShareArticle}
          onLinkedInShare={onLinkedInShareArticle}
        />
      ))}

      {articles.length > 0 && (
        <Text
          variant="tiny"
          style={{
            textAlign: 'center',
            color: '#605e5c',
            paddingTop: 8,
            opacity: 0.6
          }}
        >
          Showing {articles.length} recommendation{articles.length !== 1 ? 's' : ''} • Powered by AI
        </Text>
      )}
    </Stack>
  );
};

export default ArticleList;
