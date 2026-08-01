using System;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookQuickMove
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            DebugLog("ThisAddIn_Startup");

            // The persisted snapshot contains strings only, so it can be loaded before touching
            // any Outlook stores. The first shortcut then opens from memory instead of doing disk
            // I/O (and, when a valid snapshot exists, never performs a MAPI folder walk).
            OutlookFolderEnumerator.WarmCacheFromDisk();

            // Subscribe to Stores.StoreAdd so newly mounted data files are recorded into the
            // store-root baseline without scanning the whole Stores collection on the hot path.
            StoreRootTracker.Start(this.Application);
            OutlookShortcutManager.Start();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event reliably. Best-effort: unsubscribe and
            //    release the one retained Stores reference if it does fire.
            OutlookShortcutManager.Stop();
            StoreRootTracker.Stop();
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            DebugLog("CreateRibbonExtensibilityObject");
            return new QuickMoveRibbon();
        }

        internal static void DebugLog(string message)
        {
            QuickMoveLog.Write(message);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
