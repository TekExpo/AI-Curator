import * as React from 'react';
import { useState, useEffect } from 'react';
// Load Quill's snow theme CSS via webpack CSS loader (SPFx-compatible pattern)
// eslint-disable-next-line @typescript-eslint/no-require-imports
require('quill/dist/quill.snow.css');
import {
  Panel,
  PanelType,
  Stack,
  Text,
  PrimaryButton,
  DefaultButton,
  Label,
  Link
} from '@fluentui/react';
import ReactQuill from 'react-quill';
import { ILinkedInPanelProps } from './ILinkedInPanelProps';

const LINKEDIN_BLUE = '#0A66C2';

/** Quill toolbar configuration – basic rich-text set */
const QUILL_MODULES = {
  toolbar: [
    ['bold', 'italic', 'underline'],
    [{ list: 'ordered' }, { list: 'bullet' }],
    ['link'],
    ['clean']
  ]
};

const QUILL_FORMATS = ['bold', 'italic', 'underline', 'list', 'bullet', 'link'];

/** Escape plain text for safe inclusion in HTML */
const escapeHtml = (text: string): string =>
  text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

/** Convert a plain-text summary (with newlines) to basic HTML paragraphs */
const summaryToHtml = (text: string): string => {
  if (!text) return '';
  return text
    .split('\n')
    .map((line) => `<p>${escapeHtml(line) || '<br>'}</p>`)
    .join('');
};

/** Strip HTML tags and return plain text (for clipboard copy) */
const htmlToPlainText = (html: string): string =>
  html.replace(/<[^>]*>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').trim();

/**
 * Slide-in panel for sharing an article to LinkedIn.
 * The user edits a description, then clicks "Share on LinkedIn" which
 * opens LinkedIn's share composer pre-populated with the article title,
 * description, and URL — the user simply clicks "Post" inside LinkedIn.
 */
const LinkedInPanel: React.FC<ILinkedInPanelProps> = (props) => {
  const { isOpen, articleUrl, articleTitle, articleSummary, onDismiss } = props;

  const [description, setDescription] = useState<string>('');
  const [isSharing, setIsSharing] = useState(false);
  const [copied, setCopied] = useState(false);

  // Reset and pre-populate whenever the panel opens
  useEffect(() => {
    if (isOpen) {
      setDescription(articleSummary?.trim() ? summaryToHtml(articleSummary.trim()) : '');
      setIsSharing(false);
      setCopied(false);
    }
  }, [isOpen]);

  const handleDismiss = (): void => {
    setDescription('');
    setIsSharing(false);
    setCopied(false);
    onDismiss();
  };

  const handleCopyToClipboard = (): void => {
    const plainText = [
      htmlToPlainText(description),
      articleUrl,
      'Shared via AI Curator \u2013 Article Recommender'
    ].filter(Boolean).join('\n\n');
    if (navigator.clipboard && plainText) {
      navigator.clipboard.writeText(plainText).then(() => {
        setCopied(true);
        setTimeout(() => setCopied(false), 2000);
      }).catch(() => { /* non-fatal */ });
    }
  };
  const handleShare = (): void => {
    setIsSharing(true);

    // Strip HTML to get plain text — LinkedIn's shareArticle endpoint accepts plain text
    // in its summary param and renders it in the composer.
    const plainText = htmlToPlainText(description);

    // shareArticle?mini=true pre-populates LinkedIn's post composer with the
    // title, summary, and URL. The user just clicks "Post" in the LinkedIn window.
    const params = new URLSearchParams({
      mini: 'true',
      url: articleUrl,
      title: articleTitle,
      summary: plainText
    });
    const linkedInUrl = `https://www.linkedin.com/shareArticle?${params.toString()}`;
    window.open(linkedInUrl, '_blank', 'noopener,noreferrer');

    // Reset sharing state after a short delay so button becomes re-usable
    setTimeout(() => setIsSharing(false), 1500);
  };

  const previewHtml =
    (description.trim() || '') +
    (articleUrl
      ? `<p><a href="${escapeHtml(articleUrl)}" target="_blank" rel="noopener noreferrer">${escapeHtml(articleUrl)}</a></p>`
      : '') +
    '<p><em>Shared via AI Curator \u2013 Article Recommender</em></p>';

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={handleDismiss}
      type={PanelType.medium}
      headerText="Share on LinkedIn"
      closeButtonAriaLabel="Close"
      isFooterAtBottom
      onRenderFooterContent={() => (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text={isSharing ? 'Opening LinkedIn…' : 'Share on LinkedIn'}
            onClick={handleShare}
            disabled={isSharing}
            styles={{
              root: { backgroundColor: LINKEDIN_BLUE, borderColor: LINKEDIN_BLUE },
              rootHovered: { backgroundColor: '#004182', borderColor: '#004182' },
              rootDisabled: { backgroundColor: '#c8c6c4', borderColor: '#c8c6c4' }
            }}
            onRenderIcon={() => (
              <svg
                xmlns="http://www.w3.org/2000/svg"
                viewBox="0 0 24 24"
                width="16"
                height="16"
                aria-hidden="true"
                focusable="false"
                style={{ marginRight: 6, fill: 'currentColor' }}
              >
                <path d="M20.447 20.452h-3.554v-5.569c0-1.328-.027-3.037-1.852-3.037-1.853 0-2.136 1.445-2.136 2.939v5.667H9.351V9h3.414v1.561h.046c.477-.9 1.637-1.85 3.37-1.85 3.601 0 4.267 2.37 4.267 5.455v6.286zM5.337 7.433a2.062 2.062 0 0 1-2.063-2.065 2.064 2.064 0 1 1 2.063 2.065zm1.782 13.019H3.555V9h3.564v11.452zM22.225 0H1.771C.792 0 0 .774 0 1.729v20.542C0 23.227.792 24 1.771 24h20.451C23.2 24 24 23.227 24 22.271V1.729C24 .774 23.2 0 22.222 0h.003z" />
              </svg>
            )}
          />
          <DefaultButton text="Cancel" onClick={handleDismiss} disabled={isSharing} />
        </Stack>
      )}
    >
      <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 8 }}>

        {/* Article reference */}
        <Stack
          style={{
            padding: '10px 14px',
            borderRadius: 6,
            backgroundColor: '#e8f0fb',
            border: `1px solid #b3caf0`
          }}
        >
          <Text variant="tiny" style={{ color: '#605e5c', marginBottom: 4 }}>Sharing article</Text>
          <Text variant="medium" style={{ fontWeight: 600, color: LINKEDIN_BLUE }}>
            {articleTitle}
          </Text>
          <Text variant="tiny" style={{ color: '#605e5c', wordBreak: 'break-all', marginTop: 4, opacity: 0.8 }}>
            {articleUrl}
          </Text>
        </Stack>

        {/* Rich-text description (Quill editor) */}
        <Stack tokens={{ childrenGap: 4 }}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Label>Description</Label>
            {articleSummary?.trim() && (
              <Link
                onClick={() => setDescription(summaryToHtml(articleSummary.trim()))}
                style={{ fontSize: 12, color: LINKEDIN_BLUE }}
              >
                Reset to article summary
              </Link>
            )}
          </Stack>
          <ReactQuill
            theme="snow"
            value={description}
            onChange={(val) => setDescription(val)}
            modules={QUILL_MODULES}
            formats={QUILL_FORMATS}
            readOnly={isSharing}
            placeholder="Edit the description for your LinkedIn post…"
          />
          <Text variant="tiny" style={{ color: '#a19f9d' }}>
            Edit the description above. When you click "Share on LinkedIn", LinkedIn’s post composer will open with your title, description, and article link already filled in — just click Post.
          </Text>
        </Stack>

        {/* Live preview */}
        <Stack
          style={{
            padding: '10px 12px',
            borderRadius: 6,
            backgroundColor: '#f9f9f9',
            border: '1px solid #edebe9'
          }}
        >
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center" style={{ marginBottom: 6 }}>
            <Text variant="tiny" style={{ color: '#605e5c', fontWeight: 600 }}>Post preview</Text>
            <Link
              onClick={handleCopyToClipboard}
              style={{ fontSize: 12, color: copied ? '#107C10' : LINKEDIN_BLUE }}
            >
              {copied ? '✓ Copied!' : 'Copy to clipboard'}
            </Link>
          </Stack>
          {/* Quill output is sanitised. articleUrl is escaped before insertion. */}
          <div
            dangerouslySetInnerHTML={{ __html: previewHtml }}
            style={{ fontSize: 14, color: '#323130', lineHeight: '1.6', wordBreak: 'break-word' }}
          />
        </Stack>
      </Stack>
    </Panel>
  );
};

export default LinkedInPanel;
