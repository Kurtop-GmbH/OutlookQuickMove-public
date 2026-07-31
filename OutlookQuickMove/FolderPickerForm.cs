using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OutlookQuickMove
{
    internal sealed class FolderPickerForm : Form
    {
        private readonly FolderPickerOptions options;
        private readonly Func<FolderEnumerationResult> refreshFolders;
        private List<FolderCandidate> allFolders;
        private FolderEnumerationWarnings folderWarnings;
        private readonly Dictionary<string, FrequentTarget> frequentByKey;
        private readonly TextBox textSearch;
        private readonly ListBox listFolders;
        private readonly CheckBox checkMarkAsRead;
        private readonly Button buttonOk;
        private readonly Button buttonRefresh;
        private readonly Button buttonMove;
        private readonly Button buttonCopy;
        private readonly Button buttonGoToFolder;
        private readonly Label labelStatus;
        private readonly IntPtr anchorWindowHandle;
        private FolderPickerAction selectedAction;
        private bool closeOnDeactivateArmed;
        private bool suppressDeactivateClose;
        private static readonly Size ActionButtonSize = new Size(1, 1);
        private static readonly Size ModeButtonSize = new Size(148, 36);
        private static readonly Color HeaderBackground = Color.FromArgb(242, 242, 242);
        private static readonly Color SelectedBackground = Color.FromArgb(211, 227, 245);
        private const int HeaderRowHeight = 28;
        private const int FolderRowHeight = 27;

        public FolderPickerForm(
            IEnumerable<FolderCandidate> folders,
            FolderPickerOptions options,
            FolderEnumerationWarnings folderWarnings,
            Func<FolderEnumerationResult> refreshFolders)
        {
            this.options = options ?? FolderPickerOptions.ForQuickMove();
            this.folderWarnings = folderWarnings ?? new FolderEnumerationWarnings();
            this.refreshFolders = refreshFolders;
            allFolders = folders == null ? new List<FolderCandidate>() : folders.ToList();
            selectedAction = this.options.InitialAction;
            anchorWindowHandle = GetForegroundWindow();

            frequentByKey = new Dictionary<string, FrequentTarget>(StringComparer.OrdinalIgnoreCase);
            foreach (var target in FrequentTargetStore.LoadAll())
            {
                frequentByKey[target.Key] = target;
            }

            Text = this.options.Title;
            StartPosition = FormStartPosition.Manual;
            MinimumSize = new Size(680, 440);
            Size = new Size(820, 620);
            Font = SystemFonts.MessageBoxFont;
            FormBorderStyle = FormBorderStyle.SizableToolWindow;
            KeyPreview = true;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            KeyDown += HandleFormKeyDown;
            ShowIcon = false;
            Deactivate += HandleFormDeactivate;

            var layout = new TableLayoutPanel
            {
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Padding = new Padding(8),
                RowCount = 4
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            buttonMove = CreateModeButton("&Verschieben", FolderPickerAction.Move);
            buttonCopy = CreateModeButton("&Kopieren", FolderPickerAction.Copy);
            buttonGoToFolder = CreateModeButton("Zum &Ordner", FolderPickerAction.GoToFolder);

            var actionPanel = new TableLayoutPanel
            {
                AutoSize = false,
                ColumnCount = 3,
                Dock = DockStyle.Fill,
                Height = ModeButtonSize.Height,
                Margin = new Padding(0, 0, 0, 8),
                Padding = new Padding(0),
                RowCount = 1
            };
            actionPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.333F));
            actionPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.333F));
            actionPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.334F));
            actionPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            actionPanel.Controls.Add(buttonMove, 0, 0);
            actionPanel.Controls.Add(buttonCopy, 1, 0);
            actionPanel.Controls.Add(buttonGoToFolder, 2, 0);

            textSearch = new TextBox
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                Font = new Font(Font.FontFamily, Font.Size + 1F, FontStyle.Regular),
                Height = 34,
                Margin = new Padding(0, 0, 0, 6)
            };
            textSearch.TextChanged += delegate { ApplyFilter(); };
            textSearch.KeyDown += HandleSearchKeyDown;

            buttonRefresh = new Button
            {
                Enabled = this.refreshFolders != null,
                Margin = new Padding(8, 0, 0, 6),
                Size = new Size(120, 32),
                Text = "Aktualisieren"
            };
            buttonRefresh.Click += delegate { RefreshFolderList(); };

            var searchPanel = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                Padding = new Padding(0),
                RowCount = 1
            };
            searchPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            searchPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            searchPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            searchPanel.Controls.Add(textSearch, 0, 0);
            searchPanel.Controls.Add(buttonRefresh, 1, 0);

            listFolders = new ListBox
            {
                BackColor = SystemColors.Window,
                BorderStyle = BorderStyle.None,
                Dock = DockStyle.Fill,
                DrawMode = DrawMode.OwnerDrawVariable,
                IntegralHeight = false,
                Margin = new Padding(0, 0, 0, 8)
            };
            listFolders.DrawItem += DrawFolderRow;
            listFolders.MeasureItem += MeasureFolderRow;
            listFolders.SelectedIndexChanged += delegate { UpdateOkState(); };
            listFolders.DoubleClick += delegate { ConfirmSelection(); };
            listFolders.KeyDown += HandleListKeyDown;

            checkMarkAsRead = new CheckBox
            {
                AutoSize = true,
                Checked = MarkAsReadSetting.Load(),
                Dock = DockStyle.Right,
                Margin = new Padding(12, 4, 0, 0),
                Text = "Als gelesen verschieben"
            };

            labelStatus = new Label
            {
                AutoSize = true,
                Dock = DockStyle.Left,
                ForeColor = SystemColors.GrayText,
                Margin = new Padding(0, 4, 0, 0)
            };

            buttonOk = new Button
            {
                DialogResult = DialogResult.None,
                Enabled = false,
                TabStop = false,
                Visible = false,
                Size = ActionButtonSize,
                Text = this.options.ConfirmButtonText
            };
            buttonOk.Click += delegate { ConfirmSelection(); };

            var buttonCancel = new Button
            {
                DialogResult = DialogResult.Cancel,
                TabStop = false,
                Visible = false,
                Size = ActionButtonSize,
                Text = "Abbrechen"
            };

            var footerPanel = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                Padding = new Padding(0),
                RowCount = 1
            };
            footerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            footerPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            footerPanel.Controls.Add(labelStatus, 0, 0);
            footerPanel.Controls.Add(checkMarkAsRead, 1, 0);

            if (this.options.ShowActionSelector)
            {
                layout.Controls.Add(actionPanel, 0, 0);
            }

            layout.Controls.Add(searchPanel, 0, 1);
            layout.Controls.Add(listFolders, 0, 2);
            layout.Controls.Add(footerPanel, 0, 3);

            Controls.Add(layout);
            Controls.Add(buttonOk);
            Controls.Add(buttonCancel);

            AcceptButton = buttonOk;
            CancelButton = buttonCancel;

            Shown += delegate
            {
                PositionNearOutlook();
                Activate();
                textSearch.Focus();

                // Showing a modal VSTO form briefly shifts activation between Outlook and the
                // WinForms window. Arm focus-loss closing only after that startup sequence settles.
                var armTimer = new Timer { Interval = 300 };
                armTimer.Tick += delegate
                {
                    armTimer.Stop();
                    armTimer.Dispose();
                    closeOnDeactivateArmed = true;
                };
                armTimer.Start();
            };
            SetAction(selectedAction);
            ApplyFilter();
        }

        public string SelectedFolderEntryId { get; private set; }

        public string SelectedFolderStoreId { get; private set; }

        public FolderPickerAction SelectedAction
        {
            get { return selectedAction; }
        }

        public bool MarkAsReadBeforeMoving
        {
            get { return checkMarkAsRead.Checked; }
        }

        public FolderEnumerationWarnings FolderWarnings
        {
            get { return folderWarnings; }
        }

        private void ApplyFilter()
        {
            var terms = textSearch.Text
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(term => term.Trim())
                .Where(term => term.Length > 0)
                .ToArray();

            // Prefer strong leaf-name matches, then full-path and fuzzy matches. Frequency and
            // recency break ties so commonly used targets remain one or two keystrokes away.
            List<FolderCandidate> filtered;
            if (terms.Length == 0)
            {
                // Match Thunderbird's uncluttered recent-folder view. The moment the user types,
                // the complete cross-account index is searched.
                filtered = allFolders
                    .Where(folder => FrequencyCount(folder) > 0)
                    .OrderByDescending(FrequencyCount)
                    .ThenByDescending(FrequencyTicks)
                    .ThenBy(folder => folder.DisplayPath, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            else
            {
                filtered = allFolders
                    .Select(folder => new
                    {
                        Folder = folder,
                        MatchScore = CalculateMatchScore(folder.DisplayPath, terms)
                    })
                    .Where(candidate => candidate.MatchScore >= 0)
                    .OrderByDescending(candidate => candidate.MatchScore)
                    .ThenByDescending(candidate => FrequencyCount(candidate.Folder))
                    .ThenByDescending(candidate => FrequencyTicks(candidate.Folder))
                    .ThenBy(candidate => candidate.Folder.DisplayPath, StringComparer.OrdinalIgnoreCase)
                    .Select(candidate => candidate.Folder)
                    .ToList();
            }

            var rows = BuildRows(filtered);

            listFolders.BeginUpdate();
            try
            {
                listFolders.DataSource = null;
                listFolders.DataSource = rows;

                if (rows.Count > 1)
                {
                    listFolders.SelectedIndex = FindNextFolderIndex(-1, 1);
                }
            }
            finally
            {
                listFolders.EndUpdate();
            }

            labelStatus.Text = terms.Length == 0
                ? (filtered.Count == 0
                    ? "Tippen, um " + allFolders.Count + " Ordner zu durchsuchen"
                    : filtered.Count + " zuletzt verwendete Ordner")
                : filtered.Count + " von " + allFolders.Count + " Ordnern";
            UpdateOkState();
        }

        private static List<FolderPickerRow> BuildRows(IEnumerable<FolderCandidate> folders)
        {
            var rows = new List<FolderPickerRow>();
            foreach (var group in (folders ?? Enumerable.Empty<FolderCandidate>())
                .GroupBy(folder => folder.StoreId, StringComparer.OrdinalIgnoreCase))
            {
                var groupedFolders = group.ToList();
                if (groupedFolders.Count == 0)
                {
                    continue;
                }

                rows.Add(FolderPickerRow.Header(groupedFolders[0].AccountName));
                rows.AddRange(groupedFolders.Select(FolderPickerRow.FolderItem));
            }

            return rows;
        }

        private FrequentTarget GetFrequency(FolderCandidate folder)
        {
            FrequentTarget target;
            return frequentByKey.TryGetValue(FrequentTarget.BuildKey(folder.StoreId, folder.EntryId), out target)
                ? target
                : null;
        }

        private int FrequencyCount(FolderCandidate folder)
        {
            var target = GetFrequency(folder);
            return target == null ? 0 : target.Count;
        }

        private long FrequencyTicks(FolderCandidate folder)
        {
            var target = GetFrequency(folder);
            return target == null ? 0L : target.LastUsedUtc.Ticks;
        }

        private static int CalculateMatchScore(string displayPath, IEnumerable<string> terms)
        {
            var path = displayPath ?? string.Empty;
            var separatorIndex = path.LastIndexOf('\\');
            var leaf = separatorIndex >= 0 ? path.Substring(separatorIndex + 1) : path;
            var total = 0;
            foreach (var term in terms)
            {
                var termScore = CalculateTermScore(path, leaf, term);
                if (termScore < 0)
                {
                    return -1;
                }

                total += termScore;
            }

            return total;
        }

        private static int CalculateTermScore(string path, string leaf, string term)
        {
            if (string.Equals(leaf, term, StringComparison.OrdinalIgnoreCase))
            {
                return 1000;
            }

            if (leaf.StartsWith(term, StringComparison.OrdinalIgnoreCase))
            {
                return 800 - Math.Min(leaf.Length - term.Length, 100);
            }

            var leafIndex = leaf.IndexOf(term, StringComparison.OrdinalIgnoreCase);
            if (leafIndex >= 0)
            {
                return 600 - Math.Min(leafIndex, 100);
            }

            var pathIndex = path.IndexOf(term, StringComparison.OrdinalIgnoreCase);
            if (pathIndex >= 0)
            {
                return 350 - Math.Min(pathIndex, 100);
            }

            if (IsSubsequence(term, leaf))
            {
                return 180;
            }

            return IsSubsequence(term, path) ? 80 : -1;
        }

        private static bool IsSubsequence(string needle, string haystack)
        {
            if (string.IsNullOrEmpty(needle))
            {
                return true;
            }

            var matched = 0;
            for (int i = 0; i < haystack.Length && matched < needle.Length; i++)
            {
                if (char.ToUpperInvariant(haystack[i]) == char.ToUpperInvariant(needle[matched]))
                {
                    matched++;
                }
            }

            return matched == needle.Length;
        }

        private void UpdateOkState()
        {
            var selected = listFolders.SelectedItem as FolderPickerRow;
            buttonOk.Enabled = selected != null && selected.Folder != null;
        }

        private void RefreshFolderList()
        {
            if (refreshFolders == null)
            {
                return;
            }

            var selectedKey = GetSelectedFolderKey();
            buttonRefresh.Enabled = false;
            buttonOk.Enabled = false;
            try
            {
                FolderEnumerationResult refreshed;
                using (BusyCursor.Show())
                {
                    refreshed = refreshFolders();
                }

                if (refreshed == null || refreshed.Folders.Count == 0)
                {
                    ShowPickerMessage(
                        "Nach dem Aktualisieren wurden keine geeigneten E-Mail-Ordner gefunden.",
                        MessageBoxIcon.Warning);
                    return;
                }

                allFolders = refreshed.Folders.ToList();
                folderWarnings = refreshed.Warnings ?? new FolderEnumerationWarnings();
                QuickMoveLog.Write("folder picker refreshed: folders=" + allFolders.Count
                    + ", warnings=" + folderWarnings.Count + ".");
                ApplyFilter();
                RestoreSelection(selectedKey);
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("folder picker refresh failed.", ex);
                ShowPickerMessage(
                    "Die Ordnerliste konnte nicht aktualisiert werden.\n\n"
                    + "Details stehen im Protokoll: %TEMP%\\OutlookQuickMove.log",
                    MessageBoxIcon.Warning);
            }
            finally
            {
                buttonRefresh.Enabled = true;
                textSearch.Focus();
            }
        }

        private string GetSelectedFolderKey()
        {
            var selected = listFolders.SelectedItem as FolderPickerRow;
            return selected == null || selected.Folder == null
                ? string.Empty
                : FrequentTarget.BuildKey(selected.Folder.StoreId, selected.Folder.EntryId);
        }

        private void RestoreSelection(string folderKey)
        {
            if (string.IsNullOrEmpty(folderKey))
            {
                return;
            }

            for (int i = 0; i < listFolders.Items.Count; i++)
            {
                var row = listFolders.Items[i] as FolderPickerRow;
                var folder = row == null ? null : row.Folder;
                if (folder != null && string.Equals(
                    FrequentTarget.BuildKey(folder.StoreId, folder.EntryId),
                    folderKey,
                    StringComparison.OrdinalIgnoreCase))
                {
                    listFolders.SelectedIndex = i;
                    return;
                }
            }
        }

        private void HandleSearchKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                MoveListSelection(1);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Up)
            {
                MoveListSelection(-1);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void HandleListKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                MoveListSelection(1);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Up)
            {
                MoveListSelection(-1);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                ConfirmSelection();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void MoveListSelection(int offset)
        {
            if (listFolders.Items.Count == 0)
            {
                return;
            }

            if (listFolders.SelectedIndex < 0)
            {
                listFolders.SelectedIndex = FindNextFolderIndex(
                    offset < 0 ? listFolders.Items.Count : -1,
                    offset);
                return;
            }

            var nextIndex = FindNextFolderIndex(listFolders.SelectedIndex, offset);
            if (nextIndex >= 0)
            {
                listFolders.SelectedIndex = nextIndex;
            }
        }

        private int FindNextFolderIndex(int startIndex, int offset)
        {
            var index = startIndex + offset;
            while (index >= 0 && index < listFolders.Items.Count)
            {
                var row = listFolders.Items[index] as FolderPickerRow;
                if (row != null && row.Folder != null)
                {
                    return index;
                }

                index += offset;
            }

            return -1;
        }

        private void ConfirmSelection()
        {
            var row = listFolders.SelectedItem as FolderPickerRow;
            var candidate = row == null ? null : row.Folder;
            if (candidate == null)
            {
                return;
            }

            // Remember the checkbox state so it is restored next time, even after an
            // Outlook restart. Persisted on confirm so it reflects the last move the user made.
            if (options.ShowMarkAsRead && selectedAction == FolderPickerAction.Move)
            {
                MarkAsReadSetting.Save(checkMarkAsRead.Checked);
            }

            // Record this selection so the folder rises in the frequent-targets ordering. Go to
            // Folder reuses that ordering read-only, so it never records (keeps move data clean).
            if (options.RecordFrequentUse && selectedAction != FolderPickerAction.GoToFolder)
            {
                FrequentTargetStore.RecordUse(candidate.StoreId, candidate.EntryId, candidate.DisplayPath);
            }

            SelectedFolderEntryId = candidate.EntryId;
            SelectedFolderStoreId = candidate.StoreId;
            DialogResult = DialogResult.OK;
            Close();
        }

        private Button CreateModeButton(string text, FolderPickerAction action)
        {
            var button = new Button
            {
                FlatStyle = FlatStyle.Flat,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 6, 0),
                Text = text
            };
            button.FlatAppearance.BorderColor = Color.FromArgb(190, 190, 190);
            button.Click += delegate { SetAction(action); };
            return button;
        }

        private void SetAction(FolderPickerAction action)
        {
            selectedAction = action;
            UpdateModeButton(buttonMove, action == FolderPickerAction.Move);
            UpdateModeButton(buttonCopy, action == FolderPickerAction.Copy);
            UpdateModeButton(buttonGoToFolder, action == FolderPickerAction.GoToFolder);

            checkMarkAsRead.Visible = options.ShowMarkAsRead && action == FolderPickerAction.Move;
            switch (action)
            {
                case FolderPickerAction.Copy:
                    buttonOk.Text = "Kopieren";
                    break;
                case FolderPickerAction.GoToFolder:
                    buttonOk.Text = "Öffnen";
                    break;
                default:
                    buttonOk.Text = options.ConfirmButtonText;
                    break;
            }

            textSearch.Focus();
        }

        private static void UpdateModeButton(Button button, bool selected)
        {
            button.BackColor = selected ? Color.FromArgb(225, 225, 225) : SystemColors.Control;
            button.ForeColor = SystemColors.ControlText;
        }

        private void HandleFormKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                DialogResult = DialogResult.Cancel;
                Close();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                ConfirmSelection();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
            else if (options.ShowActionSelector && e.Control
                && (e.KeyCode == Keys.Left || e.KeyCode == Keys.Right))
            {
                var direction = e.KeyCode == Keys.Left ? -1 : 1;
                var actionIndex = (int)selectedAction + direction;
                actionIndex = Math.Max((int)FolderPickerAction.Move, Math.Min(
                    (int)FolderPickerAction.GoToFolder,
                    actionIndex));
                SetAction((FolderPickerAction)actionIndex);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        private void HandleFormDeactivate(object sender, EventArgs e)
        {
            if (!closeOnDeactivateArmed || suppressDeactivateClose || DialogResult != DialogResult.None)
            {
                return;
            }

            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void ShowPickerMessage(string message, MessageBoxIcon icon)
        {
            suppressDeactivateClose = true;
            try
            {
                MessageBox.Show(this, message, options.Title, MessageBoxButtons.OK, icon);
            }
            finally
            {
                suppressDeactivateClose = false;
            }
        }

        private void PositionNearOutlook()
        {
            try
            {
                if (anchorWindowHandle == IntPtr.Zero)
                {
                    CenterOnWorkingArea(Screen.FromPoint(Cursor.Position).WorkingArea);
                    return;
                }

                NativeRect outlookBounds;
                if (!GetWindowRect(anchorWindowHandle, out outlookBounds))
                {
                    CenterOnWorkingArea(Screen.FromHandle(anchorWindowHandle).WorkingArea);
                    return;
                }

                var workingArea = Screen.FromHandle(anchorWindowHandle).WorkingArea;
                var x = Math.Min(outlookBounds.Right - Width - 18, workingArea.Right - Width);
                var y = Math.Min(outlookBounds.Top + 72, workingArea.Bottom - Height);
                Location = new Point(
                    Math.Max(workingArea.Left, x),
                    Math.Max(workingArea.Top, y));
            }
            catch (Exception ex)
            {
                QuickMoveLog.WriteVerbose("folder picker could not be positioned near Outlook: " + ex.Message);
                CenterOnWorkingArea(Screen.FromPoint(Cursor.Position).WorkingArea);
            }
        }

        private void CenterOnWorkingArea(Rectangle workingArea)
        {
            Location = new Point(
                workingArea.Left + Math.Max(0, (workingArea.Width - Width) / 2),
                workingArea.Top + Math.Max(0, (workingArea.Height - Height) / 2));
        }

        private void MeasureFolderRow(object sender, MeasureItemEventArgs e)
        {
            var row = e.Index >= 0 && e.Index < listFolders.Items.Count
                ? listFolders.Items[e.Index] as FolderPickerRow
                : null;
            e.ItemHeight = row != null && row.IsHeader ? HeaderRowHeight : FolderRowHeight;
        }

        private void DrawFolderRow(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= listFolders.Items.Count)
            {
                return;
            }

            var row = listFolders.Items[e.Index] as FolderPickerRow;
            if (row == null)
            {
                return;
            }

            var selected = !row.IsHeader && (e.State & DrawItemState.Selected) != 0;
            using (var background = new SolidBrush(
                row.IsHeader ? HeaderBackground : selected ? SelectedBackground : listFolders.BackColor))
            {
                e.Graphics.FillRectangle(background, e.Bounds);
            }

            var iconBounds = new Rectangle(
                e.Bounds.Left + (row.IsHeader ? 8 : 24),
                e.Bounds.Top + (e.Bounds.Height - 16) / 2,
                16,
                16);
            if (row.IsHeader)
            {
                DrawAccountIcon(e.Graphics, iconBounds);
                TextRenderer.DrawText(
                    e.Graphics,
                    row.AccountName,
                    Font,
                    new Rectangle(iconBounds.Right + 7, e.Bounds.Top, e.Bounds.Width - 42, e.Bounds.Height),
                    SystemColors.ControlText,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
                return;
            }

            DrawFolderIcon(e.Graphics, iconBounds, row.Folder.FolderName);
            var textLeft = iconBounds.Right + 7;
            var parentText = row.Folder.ParentPath;
            var parentWidth = string.IsNullOrEmpty(parentText)
                ? 0
                : Math.Min(
                    e.Bounds.Width * 45 / 100,
                    TextRenderer.MeasureText(e.Graphics, parentText, Font).Width + 8);
            if (parentWidth > 0)
            {
                TextRenderer.DrawText(
                    e.Graphics,
                    parentText,
                    Font,
                    new Rectangle(e.Bounds.Right - parentWidth - 8, e.Bounds.Top, parentWidth, e.Bounds.Height),
                    SystemColors.GrayText,
                    TextFormatFlags.Right | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
            }

            TextRenderer.DrawText(
                e.Graphics,
                row.Folder.FolderName,
                Font,
                new Rectangle(
                    textLeft,
                    e.Bounds.Top,
                    Math.Max(10, e.Bounds.Right - textLeft - parentWidth - 14),
                    e.Bounds.Height),
                SystemColors.ControlText,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }

        private static void DrawAccountIcon(Graphics graphics, Rectangle bounds)
        {
            using (var pen = new Pen(Color.FromArgb(0, 120, 215), 1.5F))
            {
                var envelope = new Rectangle(bounds.Left, bounds.Top + 2, bounds.Width, bounds.Height - 5);
                graphics.DrawRectangle(pen, envelope);
                graphics.DrawLine(pen, envelope.Left, envelope.Top, envelope.Left + envelope.Width / 2, envelope.Top + 6);
                graphics.DrawLine(pen, envelope.Right, envelope.Top, envelope.Left + envelope.Width / 2, envelope.Top + 6);
            }
        }

        private static void DrawFolderIcon(Graphics graphics, Rectangle bounds, string folderName)
        {
            var name = (folderName ?? string.Empty).ToLowerInvariant();
            var isInbox = name.Contains("posteingang") || name == "inbox";
            var color = isInbox ? Color.FromArgb(0, 120, 215) : Color.FromArgb(245, 166, 0);
            using (var pen = new Pen(color, 1.5F))
            using (var fill = new SolidBrush(Color.FromArgb(24, color)))
            {
                var folder = new[]
                {
                    new Point(bounds.Left, bounds.Top + 4),
                    new Point(bounds.Left + 6, bounds.Top + 4),
                    new Point(bounds.Left + 8, bounds.Top + 6),
                    new Point(bounds.Right, bounds.Top + 6),
                    new Point(bounds.Right, bounds.Bottom - 1),
                    new Point(bounds.Left, bounds.Bottom - 1)
                };
                graphics.FillPolygon(fill, folder);
                graphics.DrawPolygon(pen, folder);
                if (isInbox)
                {
                    graphics.DrawLine(pen, bounds.Left + 3, bounds.Top + 10, bounds.Right - 3, bounds.Top + 10);
                }
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr windowHandle, out NativeRect rectangle);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
    }
}
