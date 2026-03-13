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
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Label,
  Link
} from '@fluentui/react';
import ReactQuill from 'react-quill';
import { ISharePanelProps } from './ISharePanelProps';

const GREEN = '#107C10';

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

/**
 * Slide-in panel for sharing an article to a Viva Engage group.
 */
const SharePanel: React.FC<ISharePanelProps> = (props) => {
  const {
    isOpen,
    articleUrl,
    articleTitle,
    articleSummary,
    groups,
    isLoadingGroups,
    groupLoadError,
    onRetryLoadGroups,
    onDismiss,
    onPost
  } = props;

  const [selectedGroupId, setSelectedGroupId] = useState<string>('');
  const [description, setDescription] = useState<string>('');
  const [isPosting, setIsPosting] = useState(false);
  const [errorMessage, setErrorMessage] = useState('');
  const [successMessage, setSuccessMessage] = useState('');

  // Reset and pre-populate whenever the panel opens
  useEffect(() => {
    if (isOpen) {
      setSelectedGroupId('');
      setDescription(articleSummary?.trim() ? summaryToHtml(articleSummary.trim()) : '');
      setErrorMessage('');
      setSuccessMessage('');
      setIsPosting(false);
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isOpen]);

  const handleDismiss = (): void => {
    setSelectedGroupId('');
    setDescription('');
    setErrorMessage('');
    setSuccessMessage('');
    onDismiss();
  };

  const handlePost = async (): Promise<void> => {
    if (!selectedGroupId) {
      setErrorMessage('Please select a Viva Engage community.');
      return;
    }
    // Quill empty state is '<p><br></p>' — treat as no content
    const descText = description.replace(/<[^>]*>/g, '').trim();
    if (!descText) {
      setErrorMessage('Please enter a description for the post.');
      return;
    }
    setIsPosting(true);
    setErrorMessage('');
    setSuccessMessage('');
    try {
      await onPost(selectedGroupId, description);
      setSuccessMessage('Article successfully shared to Viva Engage!');
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      setErrorMessage(msg);
    } finally {
      setIsPosting(false);
    }
  };

  const groupOptions: IDropdownOption[] = groups.map((g) => ({
    key: g.id,
    text: g.name
  }));

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
      headerText="Share to Viva Engage"
      closeButtonAriaLabel="Close"
      isFooterAtBottom
      onRenderFooterContent={() => (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text={isPosting ? 'Posting…' : 'Post to Viva Engage'}
            onClick={() => { void handlePost(); }}
            disabled={isPosting || !selectedGroupId || !!successMessage}
            styles={{
              root: { backgroundColor: GREEN, borderColor: GREEN },
              rootHovered: { backgroundColor: '#0b5e0b', borderColor: '#0b5e0b' },
              rootDisabled: { backgroundColor: '#c8c6c4', borderColor: '#c8c6c4' }
            }}
            iconProps={{ iconName: 'Share' }}
          />
          <DefaultButton text="Cancel" onClick={handleDismiss} disabled={isPosting} />
        </Stack>
      )}
    >
      <Stack tokens={{ childrenGap: 16 }} style={{ paddingTop: 8 }}>

        {/* Article reference */}
        <Stack
          style={{
            padding: '10px 14px',
            borderRadius: 6,
            backgroundColor: '#f3f9f3',
            border: `1px solid #c3e6cb`
          }}
        >
          <Text variant="tiny" style={{ color: '#605e5c', marginBottom: 4 }}>Sharing article</Text>
          <Text variant="medium" style={{ fontWeight: 600, color: GREEN }}>
            {articleTitle}
          </Text>
          <Text variant="tiny" style={{ color: '#605e5c', wordBreak: 'break-all', marginTop: 4, opacity: 0.8 }}>
            {articleUrl}
          </Text>
        </Stack>

        {/* Community selector */}
        {isLoadingGroups ? (
          <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
            <Spinner size={SpinnerSize.small} />
            <Text variant="small">Loading your Viva Engage communities…</Text>
          </Stack>
        ) : groupLoadError ? (
          <Stack tokens={{ childrenGap: 6 }}>
            <MessageBar
              messageBarType={MessageBarType.error}
              actions={
                <DefaultButton
                  text="Retry"
                  iconProps={{ iconName: 'Refresh' }}
                  onClick={onRetryLoadGroups}
                  styles={{ root: { minWidth: 70, height: 28, fontSize: 12 } }}
                />
              }
            >
              {groupLoadError}
            </MessageBar>
          </Stack>
        ) : (
          <Dropdown
            label="Community"
            placeholder={groups.length === 0 ? 'No Viva Engage communities found' : 'Select a Viva Engage community'}
            options={groupOptions}
            selectedKey={selectedGroupId || undefined}
            onChange={(_e, option) => {
              if (option) setSelectedGroupId(String(option.key));
            }}
            disabled={groups.length === 0}
            required
          />
        )}

        {/* Rich-text description (Quill editor) */}
        <Stack tokens={{ childrenGap: 4 }}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Label required>Description</Label>
            {articleSummary?.trim() && (
              <Link
                onClick={() => setDescription(summaryToHtml(articleSummary.trim()))}
                style={{ fontSize: 12, color: GREEN }}
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
            readOnly={isPosting}
            placeholder="Edit the description that will appear in your Viva Engage post…"
          />
          <Text variant="tiny" style={{ color: '#a19f9d' }}>
            The article URL and “Shared via AI Curator” attribution will be appended automatically.
          </Text>
        </Stack>

        {/* Live preview – renders the HTML the post will contain */}
        <Stack
          style={{
            padding: '10px 12px',
            borderRadius: 6,
            backgroundColor: '#f9f9f9',
            border: '1px solid #edebe9'
          }}
        >
          <Text variant="tiny" style={{ color: '#605e5c', fontWeight: 600, marginBottom: 6 }}>
            Post preview
          </Text>
          {/* Quill output is sanitised. articleUrl is escaped before insertion. */}
          <div
            // eslint-disable-next-line react/no-danger
            dangerouslySetInnerHTML={{ __html: previewHtml }}
            style={{ fontSize: 14, color: '#323130', lineHeight: '1.6', wordBreak: 'break-word' }}
          />
        </Stack>

        {/* Feedback */}
        {errorMessage && (
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline
            onDismiss={() => setErrorMessage('')}
          >
            {errorMessage}
          </MessageBar>
        )}
        {successMessage && (
          <MessageBar messageBarType={MessageBarType.success} isMultiline>
            {successMessage}
          </MessageBar>
        )}
      </Stack>
    </Panel>
  );
};

export default SharePanel;
