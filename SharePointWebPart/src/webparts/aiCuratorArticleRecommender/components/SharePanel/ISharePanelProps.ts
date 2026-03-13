import { IYammerGroup } from '../../services/VivaEngageService';

export interface ISharePanelProps {
  /** Whether the panel is open */
  isOpen: boolean;
  /** The article URL to share (pre-filled, read-only) */
  articleUrl: string;
  /** The article title (shown in the panel header) */
  articleTitle: string;
  /** The article summary — pre-populates the editable description when the panel opens */
  articleSummary: string;
  /** Available Viva Engage groups for the dropdown */
  groups: IYammerGroup[];
  /** Whether the groups list is loading */
  isLoadingGroups: boolean;
  /** Error message from a failed group-load attempt (empty = none) */
  groupLoadError: string;
  /** Called to retry loading groups after an error */
  onRetryLoadGroups: () => void;
  /** Called when the panel is dismissed */
  onDismiss: () => void;
  /** Called to post to the selected group */
  onPost: (groupId: string, description: string) => Promise<void>;
}
