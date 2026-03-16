export interface ILinkedInPanelProps {
  /** Whether the panel is open */
  isOpen: boolean;
  /** The article URL to share (pre-filled, read-only) */
  articleUrl: string;
  /** The article title (shown in the panel header area) */
  articleTitle: string;
  /** The article summary — pre-populates the editable description when the panel opens */
  articleSummary: string;
  /** Called when the panel is dismissed */
  onDismiss: () => void;
}
