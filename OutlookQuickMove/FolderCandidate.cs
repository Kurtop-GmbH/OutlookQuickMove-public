namespace OutlookQuickMove
{
    /// <summary>
    /// Describes a candidate destination folder by identity only.
    /// The folder is intentionally not held as a live COM object: the underlying
    /// <c>MAPIFolder</c> is resolved on demand via <c>Application.Session.GetFolderFromID</c>
    /// at move time. This avoids retaining hundreds of COM references for the lifetime of the
    /// dialog and tolerates the folder being renamed/moved while the dialog is open.
    /// </summary>
    internal sealed class FolderCandidate
    {
        public FolderCandidate(string displayPath, string entryId, string storeId)
        {
            DisplayPath = displayPath;
            EntryId = entryId;
            StoreId = storeId;
        }

        public string DisplayPath { get; }

        public string AccountName
        {
            get
            {
                var path = DisplayPath ?? string.Empty;
                var separatorIndex = path.IndexOf('\\');
                return separatorIndex < 0 ? path : path.Substring(0, separatorIndex);
            }
        }

        public string FolderName
        {
            get
            {
                var path = DisplayPath ?? string.Empty;
                var separatorIndex = path.LastIndexOf('\\');
                return separatorIndex < 0 ? path : path.Substring(separatorIndex + 1);
            }
        }

        public string ParentPath
        {
            get
            {
                var parts = (DisplayPath ?? string.Empty).Split('\\');
                return parts.Length <= 2
                    ? string.Empty
                    : string.Join("  ›  ", parts, 1, parts.Length - 2);
            }
        }

        /// <summary>
        /// Human-friendly account/folder path used by the picker. Keep <see cref="DisplayPath"/>
        /// unchanged for persistence and searching.
        /// </summary>
        public string DisplayText
        {
            get { return (DisplayPath ?? string.Empty).Replace("\\", "  ›  "); }
        }

        public string EntryId { get; }

        public string StoreId { get; }

        public override string ToString()
        {
            return DisplayPath;
        }
    }
}
