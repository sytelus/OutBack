using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Collections.Generic;

namespace OutBack
{
    public class ItemMover
    {
        private const int BATCH_SIZE = 20; // Process items in smaller batches
        private string logFilePath;

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

        public async void Start(Outlook.MAPIFolder sourceFolder, string storeName, bool isMoveOperation, double monthsOld)
        {
            Log($"Operation started: Source={sourceFolder.FolderPath}, Store={storeName}, Move={isMoveOperation}, MonthsOld={monthsOld}");

            Outlook.Store pstStore = null;
            Outlook.MAPIFolder destFolder = null;

            try
            {
                pstStore = GetStoreByName(storeName);
                if (pstStore == null)
                {
                    throw new Exception($"Store '{storeName}' not found");
                }

                // Get or create the destination folder
                destFolder = GetOrCreateFolder(pstStore, sourceFolder.FolderPath);

                // Get all items in the source folder
                int processedItems = 0;
                int skippedItems = 0;  // New counter for skipped items
                int skipForCast = 0;   // New counter for skipped items due to cast
                int skipForPermission = 0; // New counter for skipped items due to permission
                int skipForDate = 0; // New counter for skipped items due to date
                int retries = 0;
                int errorItems = 0;
                var sourceItems = new List<object>();

                using (ProgressForm progressForm = new ProgressForm(sourceFolder.FolderPath, destFolder.FolderPath))
                {
                    progressForm.Show();
                    Stopwatch stopwatch = Stopwatch.StartNew();
                    bool isCancelled = false;
                    string lastError = "";

                    DateTime cutoffDate = DateTime.Now.AddMonths(-(int)Math.Floor(monthsOld));

                    do
                    {
                        retries++;

                        // Refresh the total count before each retry
                        Log($"Retry {retries}: Total items count = {sourceFolder.Items.Count}");

                        for (int i = 1; i <= sourceFolder.Items.Count; i++)
                        {
                            sourceItems.Add(sourceFolder.Items[i]);
                        }

                        await Task.Run(() =>
                        {
                            foreach (var item in sourceItems)
                            {
                                if (isCancelled) break;

                                processedItems++;

                                if (processedItems % BATCH_SIZE == 0)
                                {
                                    GC.Collect();
                                    GC.WaitForPendingFinalizers();
                                }

                                try
                                {
                                    var isProcessed = moveMailItem(item, isMoveOperation, monthsOld, destFolder, ref skippedItems, ref skipForCast, ref skipForPermission, ref skipForDate, cutoffDate);
                                    if (!isProcessed)
                                    {
                                        isProcessed = moveCalItem(item, isMoveOperation, monthsOld, destFolder, ref skippedItems, ref skipForCast, ref skipForPermission, ref skipForDate, cutoffDate);
                                    }
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
                                    Log($"Error processing item {processedItems}", ex);
                                    //if (ex.Message == "Cannot move the items")
                                    //    break;
                                }
                                finally
                                {
                                    if(item != null) Marshal.ReleaseComObject(item);

                                    progressForm.Invoke(new Action(() =>
                                    {
                                        isCancelled = progressForm.UpdateProgress(
                                            processedItems,
                                            processedItems + sourceFolder.Items.Count,
                                            errorItems,
                                            skippedItems,
                                            skipForCast,
                                            skipForPermission,
                                            skipForDate,
                                            lastError,
                                            stopwatch.Elapsed,
                                            retries);
                                    }));
                                }
                            }
                        });
                    }
                    while (sourceFolder.Items.Count > 0 && retries < 1);

                    stopwatch.Stop();
                    progressForm.Close();

                    string completionMessage = $"Operation completed: {processedItems}/{processedItems + sourceFolder.Items.Count} processed, {errorItems} had errors, {skippedItems} skipped";
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
                if (destFolder != null) Marshal.ReleaseComObject(destFolder);
                if (sourceFolder != null) Marshal.ReleaseComObject(sourceFolder);
                if (pstStore != null) Marshal.ReleaseComObject(pstStore);
            }
        }

        private static bool moveMailItem(object item, bool isMoveOperation, double monthsOld, Outlook.MAPIFolder destFolder, ref int skippedItems, ref int skipForCast, ref int skipForPermission, ref int skipForDate, DateTime cutoffDate)
        {
            var mailItem = item as Outlook.MailItem;
            if (mailItem == null)
            {
                return false;
            }

            Outlook.MailItem movedItem = null;
            Outlook.MailItem copiedItem = null;

            try
            {
                if (mailItem.Permission != Outlook.OlPermission.olUnrestricted)
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

                return true;
            }
            finally
            {
                if (movedItem != null) Marshal.ReleaseComObject(movedItem);
                if (copiedItem != null) Marshal.ReleaseComObject(copiedItem);
            }

        }

        private static bool moveCalItem(object item, bool isMoveOperation, double monthsOld, Outlook.MAPIFolder destFolder, ref int skippedItems, ref int skipForCast, ref int skipForPermission, ref int skipForDate, DateTime cutoffDate)
        {
            var mailItem = item as Outlook.MeetingItem;
            if (mailItem == null)
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

                return true;
            }
            finally
            {
                if (movedItem != null) Marshal.ReleaseComObject(movedItem);
                if (copiedItem != null) Marshal.ReleaseComObject(copiedItem);
            }
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

        private Outlook.MAPIFolder GetOrCreateFolder(Outlook.Store store, string folderPath)
        {
            Outlook.Folder rootFolder = null;
            Outlook.Folder currentFolder = null;
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

                    Outlook.Folder subFolder = null;
                    try
                    {
                        subFolder = currentFolder.Folders[part] as Outlook.Folder;
                    }
                    catch
                    {
                        subFolder = currentFolder.Folders.Add(part, Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
                    }

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
    }
}
