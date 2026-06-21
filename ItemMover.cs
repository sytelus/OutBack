using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Collections.Generic;
using System.Globalization;

namespace OutBack
{
    public class ItemMover
    {
        private const int BATCH_SIZE = 20; // Process items in smaller batches
        private string logFilePath;

        private sealed class ExistingItemInfo
        {
            public string EntryId { get; set; }
            public string StoreId { get; set; }
            public DateTime LastModificationTime { get; set; }
        }

        public ItemMover()
        {
            string logDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".OutBack");
            Directory.CreateDirectory(logDirectory);
            logFilePath = Path.Combine(logDirectory, $"OutBack_Log_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
        }

        private void Log(string message, Exception ex = null)
        {
            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
            if (ex != null)
            {
                logEntry += $"\nException: {ex.GetType().FullName}" +
                           $"\nMessage: {ex.Message}" +
                           $"\nStack Trace:\n{ex.StackTrace}";
                if (ex.InnerException != null)
                {
                    logEntry += $"\nInner Exception: {ex.InnerException.GetType().FullName}" +
                               $"\nInner Message: {ex.InnerException.Message}" +
                               $"\nInner Stack Trace:\n{ex.InnerException.StackTrace}";
                }
            }
            File.AppendAllText(logFilePath, logEntry + Environment.NewLine + Environment.NewLine);
        }

        public void Start(Outlook.MAPIFolder sourceFolder, string storeName, bool isMoveOperation, double monthsOld)
        {
            Start(new List<Outlook.MAPIFolder> { sourceFolder }, storeName, isMoveOperation, monthsOld);
        }

        public void Start(IList<Outlook.MAPIFolder> sourceFolders, string storeName, bool isMoveOperation, double monthsOld)
        {
            Log($"Operation started: SourceFolders={(sourceFolders == null ? 0 : sourceFolders.Count)}, Store={storeName}, Move={isMoveOperation}, MonthsOld={monthsOld}");

            Outlook.Store pstStore = null;

            try
            {
                if (sourceFolders == null || sourceFolders.Count == 0)
                {
                    MessageBox.Show("No source folders were selected.", "No folders", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                SaveAndCloseActiveInlineResponse();

                pstStore = GetStoreByName(storeName);
                if (pstStore == null)
                {
                    throw new Exception($"Store '{storeName}' not found");
                }

                int totalItems = CountItems(sourceFolders);
                int processedItems = 0;
                int skippedItems = 0;
                int skipForCast = 0;
                int skipForPermission = 0;
                int skipForInformationRights = 0;
                int skipForDate = 0;
                int skipForExisting = 0;
                int replacedExisting = 0;
                int errorItems = 0;
                int retries = 1;

                using (ProgressForm progressForm = new ProgressForm(
                    $"{sourceFolders.Count} source folder(s)",
                    pstStore.DisplayName))
                {
                    progressForm.Show();
                    Stopwatch stopwatch = Stopwatch.StartNew();
                    bool isCancelled = false;
                    string lastError = "";

                    DateTime cutoffDate = DateTime.Now.AddMonths(-(int)Math.Floor(monthsOld));

                    Log($"Retry {retries}: Total items count = {totalItems}");

                    foreach (Outlook.MAPIFolder sourceFolder in sourceFolders)
                    {
                        if (isCancelled) break;

                        Outlook.MAPIFolder destFolder = null;
                        Outlook.Items sourceItems = null;
                        Dictionary<string, List<ExistingItemInfo>> existingDestItems = null;

                        try
                        {
                            destFolder = GetOrCreateFolder(pstStore, sourceFolder);
                            existingDestItems = BuildExistingItemIndex(destFolder);
                            sourceItems = sourceFolder.Items;
                            progressForm.SetCurrentFolders(sourceFolder.FolderPath, destFolder.FolderPath);
                            Log($"Processing folder: Source={sourceFolder.FolderPath}, Destination={destFolder.FolderPath}, Items={sourceItems.Count}");

                            for (int i = sourceItems.Count; i >= 1; i--)
                            {
                                if (isCancelled) break;

                                object item = null;
                                processedItems++;

                                try
                                {
                                    item = sourceItems[i];

                                    bool isProcessed = ProcessItem(
                                        item,
                                        sourceFolder.DefaultItemType,
                                        isMoveOperation,
                                        monthsOld,
                                        destFolder,
                                        existingDestItems,
                                        ref skippedItems,
                                        ref skipForCast,
                                        ref skipForPermission,
                                        ref skipForInformationRights,
                                        ref skipForDate,
                                        ref skipForExisting,
                                        ref replacedExisting,
                                        cutoffDate);

                                    if (!isProcessed)
                                    {
                                        skippedItems++;
                                        skipForCast++;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    lastError = $"Error: {ex.GetType().Name}\nMessage: {ex.Message}\nStack Trace:\n{ex.StackTrace}";
                                    errorItems++;
                                    Log($"Error processing item {processedItems} in {sourceFolder.FolderPath}", ex);
                                }
                                finally
                                {
                                    if (item != null) Marshal.ReleaseComObject(item);

                                    isCancelled = progressForm.UpdateProgress(
                                        processedItems,
                                        totalItems,
                                        errorItems,
                                        skippedItems,
                                        skipForCast,
                                        skipForPermission,
                                        skipForInformationRights,
                                        skipForDate,
                                        skipForExisting,
                                        replacedExisting,
                                        lastError,
                                        stopwatch.Elapsed,
                                        retries);

                                    if (processedItems % BATCH_SIZE == 0)
                                    {
                                        GC.Collect();
                                        GC.WaitForPendingFinalizers();
                                    }
                                }
                            }
                        }
                        finally
                        {
                            if (sourceItems != null) Marshal.ReleaseComObject(sourceItems);
                            if (destFolder != null) Marshal.ReleaseComObject(destFolder);
                        }
                    }

                    stopwatch.Stop();
                    progressForm.Close();

                    string completionMessage = $"Operation completed: {processedItems}/{totalItems} processed, {errorItems} had errors, {skippedItems} skipped ({skipForExisting} existing, {skipForInformationRights} information rights), {replacedExisting} replaced";
                    Log(completionMessage);
                    Log($"Total time: {stopwatch.Elapsed}");

                    MessageBox.Show($"{completionMessage}. Last error: '{lastError}'",
                        "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                string errorMessage = $"An error occurred:\nType: {ex.GetType().Name}\nMessage: {ex.Message}\nStack Trace:\n{ex.StackTrace}";
                if (ex.InnerException != null)
                {
                    errorMessage += $"\n\nInner Exception:\nType: {ex.InnerException.GetType().Name}" +
                                  $"\nMessage: {ex.InnerException.Message}" +
                                  $"\nStack Trace:\n{ex.InnerException.StackTrace}";
                }
                Log("Operation failed", ex);
                MessageBox.Show(errorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Clean up final objects
                if (sourceFolders != null)
                {
                    foreach (Outlook.MAPIFolder sourceFolder in sourceFolders)
                    {
                        if (sourceFolder != null) Marshal.ReleaseComObject(sourceFolder);
                    }
                }
                if (pstStore != null) Marshal.ReleaseComObject(pstStore);
            }
        }

        private int CountItems(IList<Outlook.MAPIFolder> sourceFolders)
        {
            int totalItems = 0;

            foreach (Outlook.MAPIFolder sourceFolder in sourceFolders)
            {
                Outlook.Items sourceItems = null;
                try
                {
                    sourceItems = sourceFolder.Items;
                    totalItems += sourceItems.Count;
                }
                finally
                {
                    if (sourceItems != null) Marshal.ReleaseComObject(sourceItems);
                }
            }

            return totalItems;
        }

        private void SaveAndCloseActiveInlineResponse()
        {
            Outlook.Explorer explorer = null;
            object inlineResponse = null;

            try
            {
                explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer == null)
                    return;

                inlineResponse = explorer.ActiveInlineResponse;
                Outlook.MailItem mailItem = null;

                if (!TryCastItem(inlineResponse, out mailItem))
                    return;

                mailItem.Save();
                mailItem.Close(Outlook.OlInspectorClose.olSave);
                Log("Saved and closed active inline response before processing.");
            }
            catch (Exception ex)
            {
                Log("Unable to save and close active inline response before processing.", ex);
            }
            finally
            {
                if (inlineResponse != null) Marshal.ReleaseComObject(inlineResponse);
                if (explorer != null) Marshal.ReleaseComObject(explorer);
            }
        }

        private static bool ProcessItem(object item, Outlook.OlItemType sourceFolderItemType, bool isMoveOperation, double monthsOld, Outlook.MAPIFolder destFolder, Dictionary<string, List<ExistingItemInfo>> existingDestItems, ref int skippedItems, ref int skipForCast, ref int skipForPermission, ref int skipForInformationRights, ref int skipForDate, ref int skipForExisting, ref int replacedExisting, DateTime cutoffDate)
        {
            bool isProcessed = false;
            string itemKey = GetItemIdentityKey(item);
            if (ShouldSkipExistingItem(item, itemKey, existingDestItems, ref skippedItems, ref skipForExisting))
                return true;

            if (sourceFolderItemType == Outlook.OlItemType.olContactItem)
            {
                isProcessed = moveContactItem(item, isMoveOperation, monthsOld, destFolder, existingDestItems, itemKey, ref skippedItems, ref skipForCast, ref skipForPermission, ref skipForDate, ref replacedExisting, cutoffDate);
                if (isProcessed)
                    return true;
            }
            else if (sourceFolderItemType == Outlook.OlItemType.olAppointmentItem)
            {
                isProcessed = moveApptItem(item, isMoveOperation, monthsOld, destFolder, existingDestItems, itemKey, ref skippedItems, ref skipForCast, ref skipForPermission, ref skipForDate, ref replacedExisting, cutoffDate);
                if (isProcessed)
                    return true;
            }

            isProcessed = moveMailItem(item, isMoveOperation, monthsOld, destFolder, existingDestItems, itemKey, ref skippedItems, ref skipForCast, ref skipForPermission, ref skipForInformationRights, ref skipForDate, ref replacedExisting, cutoffDate);
            if (!isProcessed)
            {
                isProcessed = moveCalItem(item, isMoveOperation, monthsOld, destFolder, existingDestItems, itemKey, ref skippedItems, ref skipForCast, ref skipForPermission, ref skipForDate, ref replacedExisting, cutoffDate);
            }
            if (!isProcessed)
            {
                isProcessed = moveApptItem(item, isMoveOperation, monthsOld, destFolder, existingDestItems, itemKey, ref skippedItems, ref skipForCast, ref skipForPermission, ref skipForDate, ref replacedExisting, cutoffDate);
            }
            if (!isProcessed)
            {
                isProcessed = moveContactItem(item, isMoveOperation, monthsOld, destFolder, existingDestItems, itemKey, ref skippedItems, ref skipForCast, ref skipForPermission, ref skipForDate, ref replacedExisting, cutoffDate);
            }

            return isProcessed;
        }

        private static bool TryCastItem<T>(object item, out T typedItem) where T : class
        {
            typedItem = null;
            if (item == null)
                return false;

            try
            {
                typedItem = item as T;
                return typedItem != null;
            }
            catch (InvalidCastException)
            {
                return false;
            }
            catch (COMException ex)
            {
                if ((uint)ex.ErrorCode == 0x80004002)
                    return false;

                throw;
            }
        }

        // Build once per destination folder so existing-item checks are O(1) during the item loop.
        private static Dictionary<string, List<ExistingItemInfo>> BuildExistingItemIndex(Outlook.MAPIFolder destFolder)
        {
            var existingItems = new Dictionary<string, List<ExistingItemInfo>>(StringComparer.OrdinalIgnoreCase);
            Outlook.Items items = null;

            try
            {
                items = destFolder.Items;
                int itemCount = items.Count;

                for (int i = 1; i <= itemCount; i++)
                {
                    object item = null;
                    try
                    {
                        item = items[i];
                        string itemKey = GetItemIdentityKey(item);
                        string entryId = GetItemEntryId(item);
                        if (string.IsNullOrEmpty(itemKey) || string.IsNullOrEmpty(entryId))
                            continue;

                        var itemInfo = new ExistingItemInfo
                        {
                            EntryId = entryId,
                            StoreId = destFolder.StoreID,
                            LastModificationTime = GetItemLastModificationTime(item)
                        };

                        List<ExistingItemInfo> itemInfos;
                        if (!existingItems.TryGetValue(itemKey, out itemInfos))
                        {
                            itemInfos = new List<ExistingItemInfo>();
                            existingItems[itemKey] = itemInfos;
                        }

                        itemInfos.Add(itemInfo);
                    }
                    finally
                    {
                        if (item != null) Marshal.ReleaseComObject(item);
                    }
                }
            }
            finally
            {
                if (items != null) Marshal.ReleaseComObject(items);
            }

            return existingItems;
        }

        private static bool ShouldSkipExistingItem(object item, string itemKey, Dictionary<string, List<ExistingItemInfo>> existingDestItems, ref int skippedItems, ref int skipForExisting)
        {
            List<ExistingItemInfo> existingInfos;
            if (string.IsNullOrEmpty(itemKey) ||
                existingDestItems == null ||
                !existingDestItems.TryGetValue(itemKey, out existingInfos))
            {
                return false;
            }

            DateTime sourceLastModificationTime = GetItemLastModificationTime(item);
            foreach (ExistingItemInfo existingInfo in existingInfos)
            {
                if (existingInfo.LastModificationTime >= sourceLastModificationTime)
                {
                    skippedItems++;
                    skipForExisting++;
                    return true;
                }
            }

            return false;
        }

        private static void FinishDestinationWrite(Dictionary<string, List<ExistingItemInfo>> existingDestItems, string itemKey, object item, Outlook.MAPIFolder destFolder, ref int replacedExisting)
        {
            List<ExistingItemInfo> oldItems = TakeExistingItems(itemKey, existingDestItems);
            RememberDestinationItem(existingDestItems, itemKey, item, destFolder);
            replacedExisting += DeleteExistingDestinationItems(oldItems);
        }

        private static List<ExistingItemInfo> TakeExistingItems(string itemKey, Dictionary<string, List<ExistingItemInfo>> existingDestItems)
        {
            List<ExistingItemInfo> existingInfos;
            if (string.IsNullOrEmpty(itemKey) ||
                existingDestItems == null ||
                !existingDestItems.TryGetValue(itemKey, out existingInfos))
            {
                return null;
            }

            existingDestItems.Remove(itemKey);
            return existingInfos;
        }

        private static int DeleteExistingDestinationItems(List<ExistingItemInfo> existingInfos)
        {
            if (existingInfos == null || existingInfos.Count == 0)
                return 0;

            int deletedItems = 0;
            foreach (ExistingItemInfo existingInfo in existingInfos)
            {
                object existingItem = null;
                try
                {
                    existingItem = Globals.ThisAddIn.Application.Session.GetItemFromID(existingInfo.EntryId, existingInfo.StoreId);
                    DeleteItem(existingItem);
                    deletedItems++;
                }
                finally
                {
                    if (existingItem != null) Marshal.ReleaseComObject(existingItem);
                }
            }

            return deletedItems;
        }

        private static void RememberDestinationItem(Dictionary<string, List<ExistingItemInfo>> existingDestItems, string itemKey, object item, Outlook.MAPIFolder destFolder)
        {
            if (string.IsNullOrEmpty(itemKey) || existingDestItems == null || item == null)
                return;

            string entryId = GetItemEntryId(item);
            if (string.IsNullOrEmpty(entryId))
                return;

            List<ExistingItemInfo> itemInfos;
            if (!existingDestItems.TryGetValue(itemKey, out itemInfos))
            {
                itemInfos = new List<ExistingItemInfo>();
                existingDestItems[itemKey] = itemInfos;
            }

            itemInfos.Add(new ExistingItemInfo
            {
                EntryId = entryId,
                StoreId = destFolder.StoreID,
                LastModificationTime = GetItemLastModificationTime(item)
            });
        }

        private static string GetItemIdentityKey(object item)
        {
            Outlook.MailItem mailItem = null;
            if (TryCastItem(item, out mailItem))
                return GetMailItemIdentityKey(mailItem);

            Outlook.MeetingItem meetingItem = null;
            if (TryCastItem(item, out meetingItem))
                return GetMeetingItemIdentityKey(meetingItem);

            Outlook.AppointmentItem appointmentItem = null;
            if (TryCastItem(item, out appointmentItem))
                return GetAppointmentItemIdentityKey(appointmentItem);

            Outlook.ContactItem contactItem = null;
            if (TryCastItem(item, out contactItem))
                return GetContactItemIdentityKey(contactItem);

            Outlook.DistListItem distListItem = null;
            if (TryCastItem(item, out distListItem))
                return GetDistributionListIdentityKey(distListItem);

            return string.Empty;
        }

        private static string GetMailItemIdentityKey(Outlook.MailItem mailItem)
        {
            string internetMessageId = GetInternetMessageId(mailItem.PropertyAccessor);
            if (!string.IsNullOrEmpty(internetMessageId))
                return "MAIL:" + NormalizeForKey(internetMessageId);

            return BuildIdentityKey("MAIL-FALLBACK:", 3, new[]
            {
                NormalizeForKey(SafeGetString(() => mailItem.Subject)),
                NormalizeForKey(SafeGetString(() => mailItem.SenderEmailAddress)),
                DateForKey(SafeGetDateTime(() => mailItem.SentOn)),
                DateForKey(SafeGetDateTime(() => mailItem.ReceivedTime)),
                IntForKey(SafeGetInt(() => mailItem.Size))
            });
        }

        private static string GetMeetingItemIdentityKey(Outlook.MeetingItem meetingItem)
        {
            string internetMessageId = GetInternetMessageId(meetingItem.PropertyAccessor);
            if (!string.IsNullOrEmpty(internetMessageId))
                return "MEETING:" + NormalizeForKey(internetMessageId);

            return BuildIdentityKey("MEETING-FALLBACK:", 3, new[]
            {
                NormalizeForKey(SafeGetString(() => meetingItem.Subject)),
                NormalizeForKey(SafeGetString(() => meetingItem.SenderEmailAddress)),
                DateForKey(SafeGetDateTime(() => meetingItem.ReceivedTime)),
                IntForKey(SafeGetInt(() => meetingItem.Size))
            });
        }

        private static string GetAppointmentItemIdentityKey(Outlook.AppointmentItem appointmentItem)
        {
            string globalAppointmentId = SafeGetString(() => appointmentItem.GlobalAppointmentID);
            if (!string.IsNullOrEmpty(globalAppointmentId))
                return "APPOINTMENT:" + NormalizeForKey(globalAppointmentId);

            return BuildIdentityKey("APPOINTMENT-FALLBACK:", 3, new[]
            {
                NormalizeForKey(SafeGetString(() => appointmentItem.Subject)),
                DateForKey(SafeGetDateTime(() => appointmentItem.Start)),
                DateForKey(SafeGetDateTime(() => appointmentItem.End)),
                NormalizeForKey(SafeGetString(() => appointmentItem.Location))
            });
        }

        private static string GetContactItemIdentityKey(Outlook.ContactItem contactItem)
        {
            string emailKey = BuildIdentityKey("CONTACT-EMAIL:", 1, new[]
            {
                NormalizeForKey(SafeGetString(() => contactItem.Email1Address)),
                NormalizeForKey(SafeGetString(() => contactItem.Email2Address)),
                NormalizeForKey(SafeGetString(() => contactItem.Email3Address))
            });

            if (!string.IsNullOrEmpty(emailKey))
                return emailKey;

            return BuildIdentityKey("CONTACT-FALLBACK:", 2, new[]
            {
                NormalizeForKey(SafeGetString(() => contactItem.FullName)),
                NormalizeForKey(SafeGetString(() => contactItem.CompanyName)),
                NormalizePhoneForKey(SafeGetString(() => contactItem.BusinessTelephoneNumber)),
                NormalizePhoneForKey(SafeGetString(() => contactItem.MobileTelephoneNumber))
            });
        }

        private static string GetDistributionListIdentityKey(Outlook.DistListItem distListItem)
        {
            return string.Empty;
        }

        private static string GetInternetMessageId(Outlook.PropertyAccessor propertyAccessor)
        {
            try
            {
                return GetStringProperty(
                    propertyAccessor,
                    "http://schemas.microsoft.com/mapi/proptag/0x1035001F",
                    "http://schemas.microsoft.com/mapi/proptag/0x1035001E");
            }
            finally
            {
                if (propertyAccessor != null) Marshal.ReleaseComObject(propertyAccessor);
            }
        }

        private static string GetStringProperty(Outlook.PropertyAccessor propertyAccessor, params string[] schemaNames)
        {
            foreach (string schemaName in schemaNames)
            {
                try
                {
                    object value = propertyAccessor.GetProperty(schemaName);
                    if (value != null && !string.IsNullOrWhiteSpace(value.ToString()))
                        return value.ToString();
                }
                catch
                {
                }
            }

            return string.Empty;
        }

        private static DateTime GetItemLastModificationTime(object item)
        {
            Outlook.MailItem mailItem = null;
            if (TryCastItem(item, out mailItem))
                return SafeGetDateTime(() => mailItem.LastModificationTime);

            Outlook.MeetingItem meetingItem = null;
            if (TryCastItem(item, out meetingItem))
                return SafeGetDateTime(() => meetingItem.LastModificationTime);

            Outlook.AppointmentItem appointmentItem = null;
            if (TryCastItem(item, out appointmentItem))
                return SafeGetDateTime(() => appointmentItem.LastModificationTime);

            Outlook.ContactItem contactItem = null;
            if (TryCastItem(item, out contactItem))
                return SafeGetDateTime(() => contactItem.LastModificationTime);

            Outlook.DistListItem distListItem = null;
            if (TryCastItem(item, out distListItem))
                return SafeGetDateTime(() => distListItem.LastModificationTime);

            return DateTime.MinValue;
        }

        private static string GetItemEntryId(object item)
        {
            Outlook.MailItem mailItem = null;
            if (TryCastItem(item, out mailItem))
                return SafeGetString(() => mailItem.EntryID);

            Outlook.MeetingItem meetingItem = null;
            if (TryCastItem(item, out meetingItem))
                return SafeGetString(() => meetingItem.EntryID);

            Outlook.AppointmentItem appointmentItem = null;
            if (TryCastItem(item, out appointmentItem))
                return SafeGetString(() => appointmentItem.EntryID);

            Outlook.ContactItem contactItem = null;
            if (TryCastItem(item, out contactItem))
                return SafeGetString(() => contactItem.EntryID);

            Outlook.DistListItem distListItem = null;
            if (TryCastItem(item, out distListItem))
                return SafeGetString(() => distListItem.EntryID);

            return string.Empty;
        }

        private static void DeleteItem(object item)
        {
            Outlook.MailItem mailItem = null;
            if (TryCastItem(item, out mailItem))
            {
                mailItem.Delete();
                return;
            }

            Outlook.MeetingItem meetingItem = null;
            if (TryCastItem(item, out meetingItem))
            {
                meetingItem.Delete();
                return;
            }

            Outlook.AppointmentItem appointmentItem = null;
            if (TryCastItem(item, out appointmentItem))
            {
                appointmentItem.Delete();
                return;
            }

            Outlook.ContactItem contactItem = null;
            if (TryCastItem(item, out contactItem))
            {
                contactItem.Delete();
                return;
            }

            Outlook.DistListItem distListItem = null;
            if (TryCastItem(item, out distListItem))
            {
                distListItem.Delete();
            }
        }

        private static string NormalizeForKey(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim().ToLowerInvariant();
        }

        private static string BuildIdentityKey(string prefix, int minimumNonEmptyParts, string[] parts)
        {
            int nonEmptyParts = 0;
            foreach (string part in parts)
            {
                if (!string.IsNullOrEmpty(part))
                    nonEmptyParts++;
            }

            return nonEmptyParts >= minimumNonEmptyParts ? prefix + string.Join("|", parts) : string.Empty;
        }

        private static string DateForKey(DateTime value)
        {
            return value == DateTime.MinValue ? string.Empty : value.ToUniversalTime().Ticks.ToString(CultureInfo.InvariantCulture);
        }

        private static string IntForKey(int value)
        {
            return value > 0 ? value.ToString(CultureInfo.InvariantCulture) : string.Empty;
        }

        private static string NormalizePhoneForKey(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            var digits = new System.Text.StringBuilder(value.Length);
            foreach (char c in value)
            {
                if (char.IsDigit(c))
                    digits.Append(c);
            }

            return digits.ToString();
        }

        private static string SafeGetString(Func<string> getter)
        {
            try
            {
                return getter() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private static DateTime SafeGetDateTime(Func<DateTime> getter)
        {
            try
            {
                return getter();
            }
            catch
            {
                return DateTime.MinValue;
            }
        }

        private static int SafeGetInt(Func<int> getter)
        {
            try
            {
                return getter();
            }
            catch
            {
                return 0;
            }
        }

        private static bool moveMailItem(object item, bool isMoveOperation, double monthsOld, Outlook.MAPIFolder destFolder, Dictionary<string, List<ExistingItemInfo>> existingDestItems, string itemKey, ref int skippedItems, ref int skipForCast, ref int skipForPermission, ref int skipForInformationRights, ref int skipForDate, ref int replacedExisting, DateTime cutoffDate)
        {
            Outlook.MailItem mailItem = null;
            if (!TryCastItem(item, out mailItem))
            {
                return false;
            }

            Outlook.MailItem movedItem = null;
            Outlook.MailItem copiedItem = null;

            try
            {
                Outlook.OlPermission permission = mailItem.Permission;
                if (NeedsInformationRights(mailItem, permission))
                {
                    skippedItems++;
                    skipForInformationRights++;
                    return true;
                }

                if (permission != Outlook.OlPermission.olUnrestricted)
                {
                    skippedItems++;
                    skipForPermission++;
                    return true;
                }

                if (monthsOld > 0 && mailItem.ReceivedTime > cutoffDate)
                {
                    skippedItems++;
                    skipForDate++;
                    return true;
                }

                if (isMoveOperation)
                {
                    movedItem = mailItem.Move(destFolder);
                }
                else
                {
                    copiedItem = mailItem.Copy();
                    movedItem = copiedItem.Move(destFolder);
                }

                FinishDestinationWrite(existingDestItems, itemKey, movedItem, destFolder, ref replacedExisting);
                return true;
            }
            finally
            {
                if (movedItem != null) Marshal.ReleaseComObject(movedItem);
                if (copiedItem != null) Marshal.ReleaseComObject(copiedItem);
            }

        }

        private static bool NeedsInformationRights(Outlook.MailItem mailItem, Outlook.OlPermission permission)
        {
            return permission == Outlook.OlPermission.olPermissionTemplate ||
                !string.IsNullOrEmpty(SafeGetString(() => mailItem.PermissionTemplateGuid));
        }

        private static bool moveApptItem(object item, bool isMoveOperation, double monthsOld, Outlook.MAPIFolder destFolder, Dictionary<string, List<ExistingItemInfo>> existingDestItems, string itemKey, ref int skippedItems, ref int skipForCast, ref int skipForPermission, ref int skipForDate, ref int replacedExisting, DateTime cutoffDate)
        {
            Outlook.AppointmentItem mailItem = null;
            if (!TryCastItem(item, out mailItem))
            {
                return false;
            }

            Outlook.AppointmentItem movedItem = null;
            Outlook.AppointmentItem copiedItem = null;

            try
            {
                if (isMoveOperation)
                {
                    movedItem = mailItem.Move(destFolder);
                }
                else
                {
                    copiedItem = mailItem.Copy();
                    movedItem = copiedItem.Move(destFolder);
                }

                FinishDestinationWrite(existingDestItems, itemKey, movedItem, destFolder, ref replacedExisting);
                return true;
            }
            finally
            {
                if (movedItem != null) Marshal.ReleaseComObject(movedItem);
                if (copiedItem != null) Marshal.ReleaseComObject(copiedItem);
            }
        }


        private static bool moveCalItem(object item, bool isMoveOperation, double monthsOld, Outlook.MAPIFolder destFolder, Dictionary<string, List<ExistingItemInfo>> existingDestItems, string itemKey, ref int skippedItems, ref int skipForCast, ref int skipForPermission, ref int skipForDate, ref int replacedExisting, DateTime cutoffDate)
        {
            Outlook.MeetingItem mailItem = null;
            if (!TryCastItem(item, out mailItem))
            {
                return false;
            }

            Outlook.MeetingItem movedItem = null;
            Outlook.MeetingItem copiedItem = null;

            try
            {
                if (monthsOld > 0 && mailItem.ReceivedTime > cutoffDate)
                {
                    skippedItems++;
                    skipForDate++;
                    return true;
                }

                if (isMoveOperation)
                {
                    movedItem = mailItem.Move(destFolder);
                }
                else
                {
                    copiedItem = mailItem.Copy();
                    movedItem = copiedItem.Move(destFolder);
                }

                FinishDestinationWrite(existingDestItems, itemKey, movedItem, destFolder, ref replacedExisting);
                return true;
            }
            finally
            {
                if (movedItem != null) Marshal.ReleaseComObject(movedItem);
                if (copiedItem != null) Marshal.ReleaseComObject(copiedItem);
            }
        }

        private static bool moveContactItem(object item, bool isMoveOperation, double monthsOld, Outlook.MAPIFolder destFolder, Dictionary<string, List<ExistingItemInfo>> existingDestItems, string itemKey, ref int skippedItems, ref int skipForCast, ref int skipForPermission, ref int skipForDate, ref int replacedExisting, DateTime cutoffDate)
        {
            Outlook.ContactItem contactItem = null;
            if (TryCastItem(item, out contactItem))
            {
                return moveContactItem(contactItem, isMoveOperation, destFolder, existingDestItems, itemKey, ref replacedExisting);
            }

            Outlook.DistListItem distListItem = null;
            if (TryCastItem(item, out distListItem))
            {
                return moveDistributionListItem(distListItem, isMoveOperation, destFolder, existingDestItems, itemKey, ref replacedExisting);
            }

            return false;
        }

        private static bool moveContactItem(Outlook.ContactItem contactItem, bool isMoveOperation, Outlook.MAPIFolder destFolder, Dictionary<string, List<ExistingItemInfo>> existingDestItems, string itemKey, ref int replacedExisting)
        {
            object movedItem = null;
            object copiedItem = null;

            try
            {
                if (isMoveOperation)
                {
                    movedItem = contactItem.Move(destFolder);
                }
                else
                {
                    copiedItem = contactItem.Copy();
                    Outlook.ContactItem copiedContact = null;
                    if (!TryCastItem(copiedItem, out copiedContact))
                        throw new InvalidCastException("Copied contact item could not be cast to ContactItem.");

                    movedItem = copiedContact.Move(destFolder);
                }

                FinishDestinationWrite(existingDestItems, itemKey, movedItem, destFolder, ref replacedExisting);
                return true;
            }
            finally
            {
                ReleaseComObjectIfNeeded(movedItem);
                ReleaseComObjectIfNeeded(copiedItem);
            }
        }

        private static bool moveDistributionListItem(Outlook.DistListItem distListItem, bool isMoveOperation, Outlook.MAPIFolder destFolder, Dictionary<string, List<ExistingItemInfo>> existingDestItems, string itemKey, ref int replacedExisting)
        {
            object movedItem = null;
            object copiedItem = null;

            try
            {
                if (isMoveOperation)
                {
                    movedItem = distListItem.Move(destFolder);
                }
                else
                {
                    copiedItem = distListItem.Copy();
                    Outlook.DistListItem copiedList = null;
                    if (!TryCastItem(copiedItem, out copiedList))
                        throw new InvalidCastException("Copied distribution list item could not be cast to DistListItem.");

                    movedItem = copiedList.Move(destFolder);
                }

                FinishDestinationWrite(existingDestItems, itemKey, movedItem, destFolder, ref replacedExisting);
                return true;
            }
            finally
            {
                ReleaseComObjectIfNeeded(movedItem);
                ReleaseComObjectIfNeeded(copiedItem);
            }
        }

        private static void ReleaseComObjectIfNeeded(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
                Marshal.ReleaseComObject(comObject);
        }

        private Outlook.Store GetStoreByName(string storeName)
        {
            Outlook.Stores stores = null;
            try
            {
                stores = Globals.ThisAddIn.Application.Session.Stores;
                foreach (Outlook.Store store in stores)
                {
                    if (store.DisplayName == storeName)
                    {
                        return store;
                    }
                    Marshal.ReleaseComObject(store);
                }
            }
            finally
            {
                if (stores != null)
                    Marshal.ReleaseComObject(stores);
            }
            return null;
        }

        private Outlook.MAPIFolder GetOrCreateFolder(Outlook.Store store, Outlook.MAPIFolder sourceFolder)
        {
            return GetOrCreateFolder(store, sourceFolder.FolderPath, sourceFolder.DefaultItemType);
        }

        private Outlook.MAPIFolder GetOrCreateFolder(Outlook.Store store, string folderPath, Outlook.OlItemType defaultItemType)
        {
            Outlook.Folder rootFolder = null;
            Outlook.Folder currentFolder = null;
            Outlook.OlDefaultFolders folderType = GetFolderType(defaultItemType);

            try
            {
                rootFolder = store.GetRootFolder() as Outlook.Folder;
                currentFolder = rootFolder;
                var pathParts = new List<string>(folderPath.Split('\\'));

                //reverse scan pathParts to remove any empty strings
                for (int i = pathParts.Count - 1; i >= 0; i--)
                {
                    if (string.IsNullOrEmpty(pathParts[i]) || pathParts[i].Contains("@"))
                        pathParts.RemoveAt(i);
                }

                foreach (string part in pathParts)
                {
                    Outlook.Folder subFolder = GetOrCreateTypedSubFolder(currentFolder, part, defaultItemType, folderType);

                    if (currentFolder != rootFolder)
                        Marshal.ReleaseComObject(currentFolder);
                    currentFolder = subFolder;
                }

                return currentFolder;
            }
            finally
            {
                if (rootFolder != null)
                    Marshal.ReleaseComObject(rootFolder);
            }
        }

        private Outlook.Folder GetOrCreateTypedSubFolder(Outlook.Folder parentFolder, string folderName, Outlook.OlItemType defaultItemType, Outlook.OlDefaultFolders folderType)
        {
            Outlook.Folder subFolder = TryGetSubFolder(parentFolder, folderName);
            if (subFolder == null)
                return parentFolder.Folders.Add(folderName, folderType) as Outlook.Folder;

            if (subFolder.DefaultItemType == defaultItemType)
                return subFolder;

            Log($"Destination folder type mismatch: {subFolder.FolderPath} has {subFolder.DefaultItemType}; creating typed sibling for {defaultItemType}.");
            Marshal.ReleaseComObject(subFolder);

            string typedFolderName = $"{folderName} ({GetFolderTypeLabel(defaultItemType)})";
            subFolder = TryGetSubFolder(parentFolder, typedFolderName);
            if (subFolder == null)
                return parentFolder.Folders.Add(typedFolderName, folderType) as Outlook.Folder;

            if (subFolder.DefaultItemType == defaultItemType)
                return subFolder;

            Marshal.ReleaseComObject(subFolder);
            string uniqueTypedFolderName = $"{typedFolderName} {DateTime.Now:yyyyMMddHHmmss}";
            return parentFolder.Folders.Add(uniqueTypedFolderName, folderType) as Outlook.Folder;
        }

        private static Outlook.Folder TryGetSubFolder(Outlook.Folder parentFolder, string folderName)
        {
            try
            {
                return parentFolder.Folders[folderName] as Outlook.Folder;
            }
            catch
            {
                return null;
            }
        }

        private static Outlook.OlDefaultFolders GetFolderType(Outlook.OlItemType defaultItemType)
        {
            switch (defaultItemType)
            {
                case Outlook.OlItemType.olAppointmentItem:
                    return Outlook.OlDefaultFolders.olFolderCalendar;
                case Outlook.OlItemType.olContactItem:
                    return Outlook.OlDefaultFolders.olFolderContacts;
                case Outlook.OlItemType.olTaskItem:
                    return Outlook.OlDefaultFolders.olFolderTasks;
                case Outlook.OlItemType.olNoteItem:
                    return Outlook.OlDefaultFolders.olFolderNotes;
                case Outlook.OlItemType.olJournalItem:
                    return Outlook.OlDefaultFolders.olFolderJournal;
                default:
                    return Outlook.OlDefaultFolders.olFolderInbox;
            }
        }

        private static string GetFolderTypeLabel(Outlook.OlItemType defaultItemType)
        {
            switch (defaultItemType)
            {
                case Outlook.OlItemType.olAppointmentItem:
                    return "Calendar";
                case Outlook.OlItemType.olContactItem:
                    return "Contacts";
                case Outlook.OlItemType.olTaskItem:
                    return "Tasks";
                case Outlook.OlItemType.olNoteItem:
                    return "Notes";
                case Outlook.OlItemType.olJournalItem:
                    return "Journal";
                default:
                    return "Mail";
            }
        }
    }
}
