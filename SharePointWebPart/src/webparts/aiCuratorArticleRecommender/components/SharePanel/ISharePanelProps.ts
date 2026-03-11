import { IYammerGroup } from '../../services/VivaEngageService';

export interface ISharePanelProps {
  /** Whether the panel is open */
  isOpen: boolean;
  /** The article URL to share (pre-filled, read-only) */
  articleUrl: string;
  /** The article title (shown in the panel header) */
  articleTitle: string;
  /** Available Viva Engage groups for the dropdown */
  groups: IYammerGroup[];
  /** Whether the groups list is loading */
  isLoadingGroups: boolean;
  /** Called when the panel is dismissed */
  onDismiss: () => void;
  /** Called to post to the selected group */
  onPost: (groupId: string, userComments: string) => Promise<void>;
}
