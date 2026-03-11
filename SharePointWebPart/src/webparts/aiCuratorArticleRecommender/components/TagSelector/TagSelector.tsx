import * as React from 'react';
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  DefaultButton
} from '@fluentui/react';
import { ITagSelectorProps } from './ITagSelectorProps';

const GREEN = '#107C10';
const GREEN_LIGHT = '#e8f5e9';
const CHIP_BORDER = '#c3e6cb';

const tagChipBase: React.CSSProperties = {
  display: 'inline-block',
  padding: '4px 12px',
  margin: '4px',
  borderRadius: '16px',
  fontSize: '13px',
  fontWeight: 500,
  cursor: 'pointer',
  border: '1.5px solid',
  transition: 'all 0.15s ease',
  userSelect: 'none',
  lineHeight: '1.6'
};

/**
 * Tab 1 – My Interests
 * Renders selectable tag chips and a "Save My Interests" button.
 */
const TagSelector: React.FC<ITagSelectorProps> = (props) => {
  const {
    availableTags,
    selectedTags,
    isLoading,
    errorMessage,
    successMessage,
    onTagToggle,
    onSave,
    isSaving
  } = props;

  if (isLoading) {
    return (
      <Stack horizontalAlign="center" verticalAlign="center" style={{ minHeight: 200, padding: 40 }}>
        <Spinner size={SpinnerSize.large} label="Loading available tags…" labelPosition="bottom" />
      </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 16 }} style={{ padding: '8px 0' }}>
      <Text variant="mediumPlus" style={{ fontWeight: 600, color: GREEN }}>
        Select Your Interests
      </Text>
      <Text variant="small" style={{ color: '#605e5c' }}>
        Choose the topics you are interested in. We will recommend articles matching your selection.
      </Text>

      {errorMessage && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          {errorMessage}
        </MessageBar>
      )}

      {successMessage && (
        <MessageBar messageBarType={MessageBarType.success} isMultiline>
          {successMessage}
        </MessageBar>
      )}

      {availableTags.length === 0 && !isLoading && !errorMessage && (
        <MessageBar messageBarType={MessageBarType.info}>
          No tags found in the Articles list. Please ensure the list has items with Keywords values.
        </MessageBar>
      )}

      {availableTags.length > 0 && (
        <Stack>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: '2px' }}>
            {availableTags.map((tag) => {
              const isSelected = selectedTags.includes(tag);
              return (
                <span
                  key={tag}
                  role="checkbox"
                  aria-checked={isSelected}
                  tabIndex={0}
                  onClick={() => onTagToggle(tag)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                      e.preventDefault();
                      onTagToggle(tag);
                    }
                  }}
                  style={{
                    ...tagChipBase,
                    backgroundColor: isSelected ? GREEN : '#ffffff',
                    borderColor: isSelected ? GREEN : CHIP_BORDER,
                    color: isSelected ? '#ffffff' : '#323130',
                    boxShadow: isSelected
                      ? '0 1px 4px rgba(16,124,16,0.3)'
                      : '0 1px 2px rgba(0,0,0,0.08)',
                    outline: 'none'
                  }}
                  title={isSelected ? `Deselect "${tag}"` : `Select "${tag}"`}
                >
                  {isSelected && (
                    <span style={{ marginRight: 4, fontSize: 11 }}>✓</span>
                  )}
                  {tag}
                </span>
              );
            })}
          </div>

          {selectedTags.length > 0 && (
            <Text
              variant="small"
              style={{ marginTop: 8, color: GREEN, fontWeight: 500 }}
            >
              {selectedTags.length} tag{selectedTags.length !== 1 ? 's' : ''} selected
            </Text>
          )}
        </Stack>
      )}

      <Stack horizontal tokens={{ childrenGap: 8 }} style={{ marginTop: 8 }}>
        <PrimaryButton
          text={isSaving ? 'Saving…' : 'Save My Interests'}
          onClick={onSave}
          disabled={isSaving || selectedTags.length === 0}
          iconProps={{ iconName: 'Save' }}
          styles={{
            root: { backgroundColor: GREEN, borderColor: GREEN },
            rootHovered: { backgroundColor: '#0b5e0b', borderColor: '#0b5e0b' },
            rootDisabled: { backgroundColor: '#c8c6c4', borderColor: '#c8c6c4' }
          }}
        />
        {selectedTags.length > 0 && (
          <DefaultButton
            text="Clear All"
            onClick={() => selectedTags.forEach((t) => onTagToggle(t))}
            disabled={isSaving}
          />
        )}
      </Stack>

      {selectedTags.length > 0 && (
        <Stack
          style={{
            padding: '8px 12px',
            borderRadius: 6,
            backgroundColor: GREEN_LIGHT,
            border: `1px solid ${CHIP_BORDER}`
          }}
        >
          <Text variant="small" style={{ color: '#3a3a3a' }}>
            <strong>Selected interests:</strong>{' '}
            {selectedTags.join(', ')}
          </Text>
        </Stack>
      )}
    </Stack>
  );
};

export default TagSelector;
