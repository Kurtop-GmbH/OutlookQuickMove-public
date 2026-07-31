using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace OutlookQuickMove
{
    /// <summary>
    /// Lists recent Quick Move actions with checkboxes so the user picks exactly which moved items
    /// to send back to their original folder. The most recent action's items are pre-checked (the
    /// common "undo what I just did" case). The form only collects the choice; the actual rollback
    /// and persistence are performed by the caller.
    /// </summary>
    internal sealed class UndoHistoryForm : Form
    {
        private readonly CheckedListBox listEntries;

        public UndoHistoryForm(IList<UndoEntry> entries)
        {
            var items = entries == null ? new List<UndoEntry>() : entries.ToList();

            Text = "Verschieben rückgängig";
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(640, 360);
            Size = new Size(780, 480);
            Font = SystemFonts.MessageBoxFont;
            ShowIcon = false;

            var layout = new TableLayoutPanel
            {
                ColumnCount = 1,
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                RowCount = 3
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var label = new Label
            {
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 6),
                Text = "Ausgewählte Nachrichten werden in ihre ursprünglichen Ordner zurückverschoben."
            };

            listEntries = new CheckedListBox
            {
                CheckOnClick = true,
                Dock = DockStyle.Fill,
                HorizontalScrollbar = true,
                IntegralHeight = false
            };

            // Pre-check the most recent action. Entries from one Quick Move share a BatchId; if it
            // is missing (older records), fall back to checking just the single newest entry.
            var newestBatch = items.Count > 0 ? items[0].BatchId : string.Empty;
            listEntries.BeginUpdate();
            try
            {
                for (var i = 0; i < items.Count; i++)
                {
                    var isChecked = !string.IsNullOrEmpty(newestBatch)
                        ? string.Equals(items[i].BatchId, newestBatch, StringComparison.Ordinal)
                        : i == 0;
                    listEntries.Items.Add(items[i], isChecked);
                }
            }
            finally
            {
                listEntries.EndUpdate();
            }

            layout.Controls.Add(label, 0, 0);
            layout.Controls.Add(listEntries, 0, 1);
            layout.Controls.Add(BuildButtonRow(), 0, 2);
            Controls.Add(layout);
        }

        /// <summary>The entries the user chose to undo (only set when <see cref="DialogResult.OK"/>).</summary>
        public List<UndoEntry> SelectedEntries { get; private set; }

        /// <summary>True when the user asked to clear all history instead of undoing.</summary>
        public bool ClearAllRequested { get; private set; }

        private Control BuildButtonRow()
        {
            var row = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 8, 0, 0)
            };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            var buttonSelectAll = new Button { AutoSize = true, Margin = new Padding(0, 0, 8, 0), Text = "Alle auswählen" };
            buttonSelectAll.Click += delegate { SetAllChecked(true); };

            var buttonClear = new Button { AutoSize = true, Margin = new Padding(0, 0, 8, 0), Text = "Auswahl aufheben" };
            buttonClear.Click += delegate { SetAllChecked(false); };

            var buttonClearHistory = new Button { AutoSize = true, Margin = new Padding(0), Text = "Verlauf löschen" };
            buttonClearHistory.Click += delegate { ClearHistory(); };

            var leftButtons = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                Margin = new Padding(0),
                WrapContents = false
            };
            leftButtons.Controls.Add(buttonSelectAll);
            leftButtons.Controls.Add(buttonClear);
            leftButtons.Controls.Add(buttonClearHistory);

            var buttonUndo = new Button
            {
                DialogResult = DialogResult.None,
                Margin = new Padding(8, 0, 0, 0),
                AutoSize = true,
                MinimumSize = new Size(126, 30),
                Text = "Rückgängig"
            };
            buttonUndo.Click += delegate { ConfirmUndo(); };

            var buttonCancel = new Button
            {
                DialogResult = DialogResult.Cancel,
                Margin = new Padding(8, 0, 0, 0),
                AutoSize = true,
                MinimumSize = new Size(96, 30),
                Text = "Abbrechen"
            };

            var rightButtons = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Margin = new Padding(0),
                WrapContents = false
            };
            rightButtons.Controls.Add(buttonCancel);
            rightButtons.Controls.Add(buttonUndo);

            row.Controls.Add(leftButtons, 0, 0);
            row.Controls.Add(rightButtons, 1, 0);

            AcceptButton = buttonUndo;
            CancelButton = buttonCancel;
            return row;
        }

        private void SetAllChecked(bool isChecked)
        {
            for (var i = 0; i < listEntries.Items.Count; i++)
            {
                listEntries.SetItemChecked(i, isChecked);
            }
        }

        private void ConfirmUndo()
        {
            var selected = listEntries.CheckedItems.Cast<UndoEntry>().ToList();
            if (selected.Count == 0)
            {
                MessageBox.Show(
                    "Bitte mindestens einen Verschiebevorgang auswählen.",
                    "Verschieben rückgängig",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            SelectedEntries = selected;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void ClearHistory()
        {
            var confirm = MessageBox.Show(
                "Den gesamten Rückgängig-Verlauf löschen?\n\nDabei werden keine Nachrichten verschoben.",
                "Rückgängig-Verlauf löschen",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes)
            {
                return;
            }

            ClearAllRequested = true;
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
