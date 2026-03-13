export interface ITagSelectorProps {
  /** Saved tags string from userPersonalization (comma-separated) — used to auto-select chips after search */
  savedTags: string;
  /** Success message displayed after a successful save (empty string = none) */
  successMessage: string;
  /** Called when user clicks "Save My Interests" — receives the currently selected tags */
  onSave: (selectedTags: string[]) => void;
  /** Whether the save operation is in progress */
  isSaving: boolean;
}
