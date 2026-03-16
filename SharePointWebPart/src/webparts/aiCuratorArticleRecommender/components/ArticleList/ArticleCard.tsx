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
import { IArticle } from '../../services/TopicsService';

const GREEN = '#107C10';

/** Format an ISO date string to a readable short date, e.g. "13 Mar 2026" */
function formatPublished(iso: string): string {
  try {
    return new Date(iso).toLocaleDateString(undefined, {
      day: 'numeric',
      month: 'short',
      year: 'numeric'
    });
  } catch {
    return iso;
  }
}

interface IArticleCardProps {
  article: IArticle;
  index: number;
  isSaved: boolean;
  vivaEngageEnabled: boolean;
  onSave: (article: IArticle) => Promise<void>;
  onShare: (article: IArticle) => void;
  onLinkedInShare: (article: IArticle) => void;
}

/**
 * A single article card with green-themed styling, title link,
 * description, URL, and Save / Share action buttons.
 */
const ArticleCard: React.FC<IArticleCardProps> = (props) => {
  const { article, index, isSaved, onSave, onShare, onLinkedInShare } = props;
  const [isSaving, setIsSaving] = useState(false);
  const [saved, setSaved] = useState(isSaved);
  const [linkedInHovered, setLinkedInHovered] = useState(false);

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
          {article.summary && (
            <Text
              variant="small"
              style={{ color: '#323130', lineHeight: '1.6', whiteSpace: 'pre-line' }}
            >
              {article.summary}
            </Text>
          )}

          {/* Source + Published row */}
          <Stack horizontal tokens={{ childrenGap: 12 }} style={{ flexWrap: 'wrap' }}>
            {article.source && (
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                <Icon iconName="Globe" style={{ fontSize: 12, color: '#605e5c' }} />
                <Text variant="tiny" style={{ color: '#605e5c', fontWeight: 600 }}>
                  {article.source}
                </Text>
              </Stack>
            )}
            {article.published && (
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }}>
                <Icon iconName="Calendar" style={{ fontSize: 12, color: '#605e5c' }} />
                <Text variant="tiny" style={{ color: '#605e5c' }}>
                  {formatPublished(article.published)}
                </Text>
              </Stack>
            )}
          </Stack>

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

          <TooltipHost content="Share on LinkedIn">
            <button
              aria-label="Share on LinkedIn"
              onClick={() => onLinkedInShare(article)}
              onMouseEnter={() => setLinkedInHovered(true)}
              onMouseLeave={() => setLinkedInHovered(false)}
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                justifyContent: 'center',
                width: 32,
                height: 32,
                padding: 0,
                border: 'none',
                background: 'transparent',
                cursor: 'pointer',
                color: linkedInHovered ? '#0A66C2' : '#605e5c',
                borderRadius: 2,
                flexShrink: 0
              }}
            >
              <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="16" height="16" aria-hidden="true" focusable="false">
                <path fill="currentColor" d="M20.447 20.452h-3.554v-5.569c0-1.328-.027-3.037-1.852-3.037-1.853 0-2.136 1.445-2.136 2.939v5.667H9.351V9h3.414v1.561h.046c.477-.9 1.637-1.85 3.37-1.85 3.601 0 4.267 2.37 4.267 5.455v6.286zM5.337 7.433a2.062 2.062 0 0 1-2.063-2.065 2.064 2.064 0 1 1 2.063 2.065zm1.782 13.019H3.555V9h3.564v11.452zM22.225 0H1.771C.792 0 0 .774 0 1.729v20.542C0 23.227.792 24 1.771 24h20.451C23.2 24 24 23.227 24 22.271V1.729C24 .774 23.2 0 22.222 0h.003z"/>
              </svg>
            </button>
          </TooltipHost>
        </Stack>
      </Stack>
    </Stack>
  );
};

export default ArticleCard;
