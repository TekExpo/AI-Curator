export interface ITagSelectorProps {
  /** All tags available from the Articles list */
  availableTags: string[];
  /** Currently selected tags */
  selectedTags: string[];
  /** Whether the tag list is still loading */
  isLoading: boolean;
  /** Error message to display (empty string = no error) */
  errorMessage: string;
  /** Success message to display after save (empty string = none) */
  successMessage: string;
  /** Called when user toggles a tag chip */
  onTagToggle: (tag: string) => void;
  /** Called when user clicks "Save My Interests" */
  onSave: () => void;
  /** Whether the save operation is in progress */
  isSaving: boolean;
}
