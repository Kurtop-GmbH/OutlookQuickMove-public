namespace OutlookQuickMove
{
    /// <summary>
    /// Behavior options for <see cref="FolderPickerForm"/>, so the same folder picker serves both
    /// Quick Move (move the selected mail into the chosen folder) and Go to Folder (navigate the
    /// active explorer to the chosen folder). Both modes order results by the Quick Move
    /// frequent-targets; only Quick Move records a use and persists the "Mark as read" choice.
    /// </summary>
    internal sealed class FolderPickerOptions
    {
        private FolderPickerOptions()
        {
        }

        /// <summary>Dialog title.</summary>
        public string Title { get; private set; }

        /// <summary>Text on the confirm (default) button.</summary>
        public string ConfirmButtonText { get; private set; }

        /// <summary>Whether the Move / Copy / Go to Folder action bar is shown.</summary>
        public bool ShowActionSelector { get; private set; }

        /// <summary>Action selected when the dialog opens.</summary>
        public FolderPickerAction InitialAction { get; private set; }

        /// <summary>Whether to show (and persist) the "Mark as read before moving" checkbox.</summary>
        public bool ShowMarkAsRead { get; private set; }

        /// <summary>Whether confirming records the folder into the Quick Move frequent-targets.</summary>
        public bool RecordFrequentUse { get; private set; }

        /// <summary>Options for the Quick Move dialog (unchanged existing behavior).</summary>
        public static FolderPickerOptions ForQuickMove()
        {
            return new FolderPickerOptions
            {
                Title = "Schnelles Nachrichtenverschieben",
                ConfirmButtonText = "Verschieben",
                ShowActionSelector = true,
                InitialAction = FolderPickerAction.Move,
                ShowMarkAsRead = true,
                RecordFrequentUse = true
            };
        }

        /// <summary>
        /// Options for the Go to Folder dialog: no read-state controls and no frequent-target
        /// recording (the frequent ordering is reused read-only), so jumping never alters move data.
        /// </summary>
        public static FolderPickerOptions ForGoToFolder()
        {
            return new FolderPickerOptions
            {
                Title = "Zum Ordner",
                ConfirmButtonText = "Öffnen",
                ShowActionSelector = false,
                InitialAction = FolderPickerAction.GoToFolder,
                ShowMarkAsRead = false,
                RecordFrequentUse = false
            };
        }
    }
}
