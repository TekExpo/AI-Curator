import * as React from 'react';
import { useState } from 'react';
import {
  Stack,
  Text,
  Link,
  Icon,
  IconButton,
  TooltipHost
} from '@fluentui/react';
import { IArticle } from '../../services/OpenAIService';

const GREEN = '#107C10';

interface IArticleCardProps {
  article: IArticle;
  index: number;
  isSaved: boolean;
  vivaEngageEnabled: boolean;
  onSave: (article: IArticle) => Promise<void>;
  onShare: (article: IArticle) => void;
}

/**
 * A single article card with green-themed styling, title link,
 * description, URL, and Save / Share action buttons.
 */
const ArticleCard: React.FC<IArticleCardProps> = (props) => {
  const { article, index, isSaved, vivaEngageEnabled, onSave, onShare } = props;
  const [isSaving, setIsSaving] = useState(false);
  const [saved, setSaved] = useState(isSaved);

  const handleSave = async (): Promise<void> => {
    if (saved || isSaving) return;
    setIsSaving(true);
    try {
      await onSave(article);
      setSaved(true);
    } finally {
      setIsSaving(false);
    }
  };

  const displayUrl = article.url.length > 80
    ? article.url.substring(0, 77) + '…'
    : article.url;

  return (
    <Stack
      key={`article-${index}-${article.url}`}
      style={{
        padding: '12px 16px',
        borderRadius: '8px',
        boxShadow: '0 1px 3px rgba(0,0,0,0.12), 0 1px 2px rgba(0,0,0,0.06)',
        transition: 'box-shadow 0.2s ease, transform 0.15s ease',
        backgroundColor: '#ffffff',
        border: `1px solid #edebe9`,
        marginBottom: 0
      }}
    >
      <Stack horizontal verticalAlign="start" tokens={{ childrenGap: 10 }}>
        <Icon
          iconName="TextDocument"
          style={{ fontSize: 20, color: GREEN, marginTop: 2, flexShrink: 0 }}
        />
        <Stack tokens={{ childrenGap: 4 }} grow>
          <Link
            href={article.url}
            target="_blank"
            rel="noopener noreferrer"
            style={{
              fontSize: 15,
              fontWeight: 600,
              lineHeight: '1.3',
              color: GREEN,
              textDecoration: 'none'
            }}
            title={`Open "${article.title}" in a new tab`}
          >
            {article.title}
          </Link>
          {article.description && (
            <Text variant="small" style={{ color: '#323130', lineHeight: '1.45' }}>
              {article.description}
            </Text>
          )}
          <Text variant="tiny" style={{ color: '#605e5c', wordBreak: 'break-all', opacity: 0.7 }}>
            {displayUrl}
          </Text>
        </Stack>

        {/* Action buttons */}
        <Stack horizontal tokens={{ childrenGap: 4 }} style={{ flexShrink: 0, alignSelf: 'flex-start' }}>
          <TooltipHost content={saved ? 'Already saved' : 'Save to My Interests'}>
            <IconButton
              iconProps={{ iconName: saved ? 'CheckMark' : 'Save' }}
              ariaLabel={saved ? 'Already saved' : 'Save article'}
              disabled={saved || isSaving}
              onClick={() => { void handleSave(); }}
              styles={{
                root: { color: saved ? GREEN : '#605e5c' },
                rootDisabled: { color: GREEN },
                icon: { fontSize: 16 }
              }}
            />
          </TooltipHost>

          {vivaEngageEnabled && (
            <TooltipHost content="Share to Viva Engage">
              <IconButton
                iconProps={{ iconName: 'Share' }}
                ariaLabel="Share to Viva Engage"
                onClick={() => onShare(article)}
                styles={{
                  root: { color: '#605e5c' },
                  rootHovered: { color: GREEN },
                  icon: { fontSize: 16 }
                }}
              />
            </TooltipHost>
          )}
        </Stack>
      </Stack>
    </Stack>
  );
};

export default ArticleCard;
