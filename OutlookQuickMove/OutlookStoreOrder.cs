using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookQuickMove
{
    /// <summary>
    /// Captures Outlook's own account order without walking every mounted store. Delivery stores
    /// belonging to configured accounts are ranked first; folders from additional/shared stores
    /// follow afterwards. The snapshot contains strings only and holds no COM objects.
    /// </summary>
    internal sealed class OutlookStoreOrder
    {
        private readonly Dictionary<string, int> ownStorePriorities;

        private OutlookStoreOrder(Dictionary<string, int> ownStorePriorities)
        {
            this.ownStorePriorities = ownStorePriorities
                ?? new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        }

        public static OutlookStoreOrder Capture(Outlook.Application application)
        {
            var priorities = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            Outlook.NameSpace session = null;
            Outlook.Store defaultStore = null;
            Outlook.Accounts accounts = null;
            try
            {
                if (application == null)
                {
                    return new OutlookStoreOrder(priorities);
                }

                session = application.Session;
                try
                {
                    defaultStore = session.DefaultStore;
                    AddStore(priorities, defaultStore, 0);
                }
                catch (Exception ex)
                {
                    QuickMoveLog.WriteVerbose("default store could not be ranked: " + ex.Message);
                }

                accounts = session.Accounts;
                for (var index = 1; index <= accounts.Count; index++)
                {
                    Outlook.Account account = null;
                    Outlook.Store deliveryStore = null;
                    try
                    {
                        account = accounts[index];
                        deliveryStore = account.DeliveryStore;
                        AddStore(priorities, deliveryStore, index);
                    }
                    catch (Exception ex)
                    {
                        QuickMoveLog.WriteVerbose("Outlook account " + index + " could not be ranked: " + ex.Message);
                    }
                    finally
                    {
                        ComUtil.Release(deliveryStore);
                        ComUtil.Release(account);
                    }
                }
            }
            catch (Exception ex)
            {
                QuickMoveLog.Write("failed to capture Outlook account ordering.", ex);
            }
            finally
            {
                ComUtil.Release(accounts);
                ComUtil.Release(defaultStore);
                ComUtil.Release(session);
            }

            QuickMoveLog.WriteVerbose("captured Outlook account ordering: ownStores=" + priorities.Count + ".");
            return new OutlookStoreOrder(priorities);
        }

        public int GetPriority(FolderCandidate folder)
        {
            if (folder == null)
            {
                return int.MaxValue;
            }

            int priority;
            if (ownStorePriorities.TryGetValue(folder.StoreId ?? string.Empty, out priority))
            {
                return priority;
            }

            // Some providers expose an account without a usable DeliveryStore. In the current
            // Outlook profile, personal mailboxes are displayed as addresses while shared stores
            // use friendly names such as "Admin" or "IT", so keep those address-like stores ahead
            // of shared mailboxes as a conservative fallback.
            return (folder.AccountName ?? string.Empty).IndexOf('@') >= 0 ? 500 : 1000;
        }

        private static void AddStore(Dictionary<string, int> priorities, Outlook.Store store, int priority)
        {
            if (store == null)
            {
                return;
            }

            var storeId = store.StoreID;
            if (string.IsNullOrEmpty(storeId) || priorities.ContainsKey(storeId))
            {
                return;
            }

            priorities[storeId] = priority;
        }
    }
}
