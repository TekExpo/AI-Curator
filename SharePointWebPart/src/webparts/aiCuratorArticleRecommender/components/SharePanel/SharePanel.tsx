import * as React from 'react';
import { useState } from 'react';
import {
  Panel,
  PanelType,
  Stack,
  Text,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Label
} from '@fluentui/react';
import { ISharePanelProps } from './ISharePanelProps';

const GREEN = '#107C10';

/**
 * Slide-in panel for sharing an article to a Viva Engage group.
 */
const SharePanel: React.FC<ISharePanelProps> = (props) => {
  const {
    isOpen,
    articleUrl,
    articleTitle,
    groups,
    isLoadingGroups,
    onDismiss,
    onPost
  } = props;

  const [selectedGroupId, setSelectedGroupId] = useState<string>('');
  const [userComments, setUserComments] = useState<string>('');
  const [isPosting, setIsPosting] = useState(false);
  const [errorMessage, setErrorMessage] = useState('');
  const [successMessage, setSuccessMessage] = useState('');

  const handleDismiss = (): void => {
    setSelectedGroupId('');
    setUserComments('');
    setErrorMessage('');
    setSuccessMessage('');
    onDismiss();
  };

  const handlePost = async (): Promise<void> => {
    if (!selectedGroupId) {
      setErrorMessage('Please select a Viva Engage group.');
      return;
    }
    setIsPosting(true);
    setErrorMessage('');
    setSuccessMessage('');
    try {
      await onPost(selectedGroupId, userComments);
      setSuccessMessage('Article successfully shared to Viva Engage!');
      // Do not auto-close; let user read the confirmation
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
        <Stack>
          <Label>Article</Label>
          <Text variant="medium" style={{ fontWeight: 600, color: GREEN }}>
            {articleTitle}
          </Text>
        </Stack>

        {/* Article URL (read-only) */}
        <TextField
          label="Article URL"
          value={articleUrl}
          readOnly
          styles={{ fieldGroup: { backgroundColor: '#f3f2f1' } }}
        />

        {/* Group selector */}
        {isLoadingGroups ? (
          <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
            <Spinner size={SpinnerSize.small} />
            <Text variant="small">Loading groups…</Text>
          </Stack>
        ) : (
          <Dropdown
            label="Viva Engage Group"
            placeholder="Select a group"
            options={groupOptions}
            selectedKey={selectedGroupId || undefined}
            onChange={(_e, option) => {
              if (option) setSelectedGroupId(String(option.key));
            }}
            required
          />
        )}

        {/* User comments */}
        <TextField
          label="Add a comment (optional)"
          multiline
          rows={4}
          value={userComments}
          onChange={(_e, val) => setUserComments(val ?? '')}
          placeholder="Share your thoughts about this article…"
          disabled={isPosting}
        />

        {/* Preview */}
        <Stack
          style={{
            padding: '8px 12px',
            borderRadius: 6,
            backgroundColor: '#f9f9f9',
            border: '1px solid #edebe9'
          }}
        >
          <Text variant="tiny" style={{ color: '#605e5c' }}>
            <strong>Post preview:</strong>
          </Text>
          <Text variant="small" style={{ marginTop: 4, whiteSpace: 'pre-wrap' }}>
            {[userComments?.trim(), articleUrl, 'Shared via AI Curator – Article Recommender']
              .filter((line) => line.length > 0)
              .join('\n\n')}
          </Text>
        </Stack>

        {/* Feedback messages */}
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
