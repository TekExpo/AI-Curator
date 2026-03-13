import * as React from 'react';
import { useState, useRef, useMemo, useEffect } from 'react';
import {
  Stack,
  Text,
  TextField,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Separator
} from '@fluentui/react';
import { ITagSelectorProps } from './ITagSelectorProps';
import { TopicsService } from '../../services/TopicsService';

const GREEN = '#107C10';
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
 * Renders a search box to discover topics via the suggest-topics API,
 * then renders selectable tag chips and a "Save My Interests" button.
 */
const TagSelector: React.FC<ITagSelectorProps> = (props) => {
  const { savedTags, successMessage, onSave, isSaving } = props;

  // Parse savedTags into an array for auto-selection matching
  const savedTagsArray = useMemo(
    () =>
      savedTags
        .split(',')
        .map((t) => t.trim())
        .filter((t) => t.length > 0),
    [savedTags]
  );

  const [searchQuery, setSearchQuery] = useState('');
  const [searchResults, setSearchResults] = useState<string[]>([]);
  const [selectedTags, setSelectedTags] = useState<string[]>([]);
  const [isSearching, setIsSearching] = useState(false);
  const [searchError, setSearchError] = useState(false);
  const [hasSearched, setHasSearched] = useState(false);

  const topicsService = useRef(new TopicsService());
  const initializedRef = useRef(false);

  // Seed selectedTags from saved data once it first arrives (e.g. after SP load)
  useEffect(() => {
    if (!initializedRef.current && savedTagsArray.length > 0) {
      initializedRef.current = true;
      setSelectedTags(savedTagsArray);
    }
  }, [savedTagsArray]);

  const handleSearch = async (): Promise<void> => {
    if (!searchQuery.trim()) return;
    setIsSearching(true);
    setSearchError(false);
    try {
      const topics = await topicsService.current.getSuggestedTopics(searchQuery.trim());
      setSearchResults(topics);
      setHasSearched(true);
      // Auto-select any returned topic that matches a previously saved tag
      setSelectedTags((prev) => {
        const combined = [...prev];
        topics.forEach((topic) => {
          if (
            savedTagsArray.some((s) => s.toLowerCase() === topic.toLowerCase()) &&
            !combined.some((t) => t.toLowerCase() === topic.toLowerCase())
          ) {
            combined.push(topic);
          }
        });
        return combined;
      });
    } catch {
      setSearchError(true);
      setSearchResults([]);
      setHasSearched(true);
    } finally {
      setIsSearching(false);
    }
  };

  const handleChipToggle = (tag: string): void => {
    setSelectedTags((prev) =>
      prev.indexOf(tag) !== -1 ? prev.filter((t) => t !== tag) : [...prev, tag]
    );
  };

  return (
    <Stack tokens={{ childrenGap: 16 }} style={{ padding: '8px 0' }}>
      <Text variant="mediumPlus" style={{ fontWeight: 600, color: GREEN }}>
        Select Your Interests
      </Text>
      <Text variant="small" style={{ color: '#605e5c' }}>
        Choose the topics you are interested in. We will recommend articles matching your
        selection.
      </Text>

      {/* ── My Current Interests ──────────────────────────────────────────── */}
      {savedTagsArray.length > 0 && (
        <Stack
          tokens={{ childrenGap: 8 }}
          style={{
            padding: '12px 14px',
            borderRadius: 8,
            backgroundColor: '#f3f9f3',
            border: `1px solid ${CHIP_BORDER}`
          }}
        >
          <Text variant="small" style={{ fontWeight: 600, color: GREEN }}>
            My Current Interests ({savedTagsArray.length})
          </Text>
          <Text variant="tiny" style={{ color: '#605e5c' }}>
            Click × to remove a topic. Strikethrough means it will be removed when you save.
          </Text>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: '4px', marginTop: 4 }}>
            {savedTagsArray.map((tag) => {
              const isActive = selectedTags.indexOf(tag) !== -1;
              return (
                <span
                  key={`saved-${tag}`}
                  style={{
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: 5,
                    padding: '3px 8px 3px 12px',
                    borderRadius: 16,
                    fontSize: 13,
                    fontWeight: 500,
                    border: '1.5px solid',
                    backgroundColor: isActive ? '#e8f5e9' : '#f3f2f1',
                    borderColor: isActive ? GREEN : '#c8c6c4',
                    color: isActive ? GREEN : '#a19f9d',
                    textDecoration: isActive ? 'none' : 'line-through',
                    opacity: isActive ? 1 : 0.75,
                    transition: 'all 0.15s ease'
                  }}
                >
                  {tag}
                  <button
                    onClick={() =>
                      setSelectedTags((prev) =>
                        isActive
                          ? prev.filter((t) => t !== tag)
                          : [...prev, tag]
                      )
                    }
                    title={isActive ? `Remove "${tag}" from interests` : `Restore "${tag}"`}
                    style={{
                      background: 'none',
                      border: 'none',
                      cursor: 'pointer',
                      padding: '0 2px',
                      lineHeight: 1,
                      fontSize: 16,
                      fontWeight: 700,
                      color: isActive ? '#a4262c' : GREEN,
                      display: 'flex',
                      alignItems: 'center'
                    }}
                  >
                    {isActive ? '\u00d7' : '+'}
                  </button>
                </span>
              );
            })}
          </div>
        </Stack>
      )}

      <Separator />

      {/* ── Search for new topics ─────────────────────────────────────────── */}
      <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
        <Stack.Item grow>
          <TextField
            placeholder="Search for topics..."
            value={searchQuery}
            onChange={(_, val) => setSearchQuery(val ?? '')}
            onKeyDown={(e) => {
              if (e.key === 'Enter') {
                void handleSearch();
              }
            }}
          />
        </Stack.Item>
        <PrimaryButton
          text="Search"
          onClick={() => { void handleSearch(); }}
          disabled={isSearching || !searchQuery.trim()}
          styles={{
            root: { backgroundColor: GREEN, borderColor: GREEN, minWidth: 80 },
            rootHovered: { backgroundColor: '#0b5e0b', borderColor: '#0b5e0b' },
            rootDisabled: { backgroundColor: '#c8c6c4', borderColor: '#c8c6c4' }
          }}
        />
      </Stack>

      {/* Chip area */}
      {isSearching && (
        <Stack horizontalAlign="center" style={{ padding: '20px 0' }}>
          <Spinner size={SpinnerSize.medium} label="Searching topics…" labelPosition="right" />
        </Stack>
      )}

      {!isSearching && searchError && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline>
          Unable to fetch topics. Please try again later.
        </MessageBar>
      )}

      {!isSearching && hasSearched && !searchError && searchResults.length === 0 && (
        <MessageBar messageBarType={MessageBarType.info}>
          No topics found for your search. Try different keywords.
        </MessageBar>
      )}

      {!isSearching && searchResults.length > 0 && (
        <Stack>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: '2px' }}>
            {searchResults.map((tag) => {
              const isSelected = selectedTags.indexOf(tag) !== -1;
              return (
                <span
                  key={tag}
                  role="checkbox"
                  aria-checked={isSelected}
                  tabIndex={0}
                  onClick={() => handleChipToggle(tag)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                      e.preventDefault();
                      handleChipToggle(tag);
                    }
                  }}
                  style={{
                    ...tagChipBase,
                    backgroundColor: '#ffffff',
                    borderColor: isSelected ? GREEN : CHIP_BORDER,
                    color: isSelected ? GREEN : '#323130',
                    boxShadow: isSelected
                      ? '0 1px 4px rgba(16,124,16,0.3)'
                      : '0 1px 2px rgba(0,0,0,0.08)',
                    outline: 'none'
                  }}
                  title={isSelected ? `Deselect "${tag}"` : `Select "${tag}"`}
                >
                  {tag}
                </span>
              );
            })}
          </div>

          {selectedTags.filter((t) => searchResults.indexOf(t) !== -1).length > 0 && (
            <Text variant="small" style={{ marginTop: 8, color: GREEN, fontWeight: 500 }}>
              {selectedTags.filter((t) => searchResults.indexOf(t) !== -1).length} topic
              {selectedTags.filter((t) => searchResults.indexOf(t) !== -1).length !== 1 ? 's' : ''}{' '}
              selected
            </Text>
          )}
        </Stack>
      )}

      {successMessage && (
        <MessageBar messageBarType={MessageBarType.success} isMultiline>
          {successMessage}
        </MessageBar>
      )}

      <PrimaryButton
        text={isSaving ? 'Saving…' : 'Save My Interests'}
        iconProps={{ iconName: 'Save' }}
        onClick={() => onSave(selectedTags)}
        disabled={isSaving || selectedTags.length === 0}
        styles={{
          root: { backgroundColor: GREEN, borderColor: GREEN },
          rootHovered: { backgroundColor: '#0b5e0b', borderColor: '#0b5e0b' },
          rootDisabled: { backgroundColor: '#c8c6c4', borderColor: '#c8c6c4' }
        }}
      />

      {selectedTags.length > 0 && (
        <Stack
          style={{
            padding: '8px 12px',
            borderRadius: 6,
            backgroundColor: '#e8f5e9',
            border: `1px solid ${CHIP_BORDER}`
          }}
        >
          <Text variant="small" style={{ color: '#3a3a3a' }}>
            <strong>Selected interests:</strong> {selectedTags.join(', ')}
          </Text>
        </Stack>
      )}
    </Stack>
  );
};

export default TagSelector;
