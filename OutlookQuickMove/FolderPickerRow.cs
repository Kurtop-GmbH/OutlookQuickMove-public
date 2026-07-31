namespace OutlookQuickMove
{
    /// <summary>
    /// One visual row in the Thunderbird-style picker. Account headers are deliberately separate
    /// from selectable folder rows so arrow navigation can skip them.
    /// </summary>
    internal sealed class FolderPickerRow
    {
        private FolderPickerRow(string accountName, FolderCandidate folder)
        {
            AccountName = accountName ?? string.Empty;
            Folder = folder;
        }

        public string AccountName { get; }

        public FolderCandidate Folder { get; }

        public bool IsHeader
        {
            get { return Folder == null; }
        }

        public static FolderPickerRow Header(string accountName)
        {
            return new FolderPickerRow(accountName, null);
        }

        public static FolderPickerRow FolderItem(FolderCandidate folder)
        {
            return new FolderPickerRow(folder == null ? string.Empty : folder.AccountName, folder);
        }

        public override string ToString()
        {
            if (IsHeader)
            {
                return AccountName;
            }

            if (Folder == null)
            {
                return string.Empty;
            }

            return string.IsNullOrWhiteSpace(Folder.ParentPath)
                ? Folder.FolderName
                : Folder.FolderName + ", " + Folder.ParentPath;
        }
    }
}
