using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;

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

        private void Log(string message)
        {
            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
            File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
        }

        public async void Start(Outlook.MAPIFolder sourceFolder, string storeName, bool isMoveOperation, double monthsOld)
        {
            Log($"Operation started: Source={sourceFolder.FolderPath}, Store={storeName}, Move={isMoveOperation}, MonthsOld={monthsOld}");

            Outlook.Store pstStore = null;
            Outlook.MAPIFolder destFolder = null;
            Outlook.Items items = null;

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
                items = sourceFolder.Items;
                int totalItems = items.Count;
                int processedItems = 0;
                int errorItems = 0;
                int skippedItems = 0;  // New counter for skipped items

                using (ProgressForm progressForm = new ProgressForm(sourceFolder.FolderPath, destFolder.FolderPath))
                {
                    progressForm.Show();
                    Stopwatch stopwatch = Stopwatch.StartNew();
                    bool isCancelled = false;
                    string lastError = "";

                    DateTime cutoffDate = DateTime.Now.AddMonths(-(int)Math.Floor(monthsOld));

                    await Task.Run(() =>
                    {
                        for (int i = 1; i <= totalItems; i++)
                        {
                            if (isCancelled) break;

                            if ((i % BATCH_SIZE == 0) || (i == totalItems))
                            {
                                // Force garbage collection periodically
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                            }

                            Outlook.MailItem item = null;
                            Outlook.MailItem movedItem = null;
                            Outlook.MailItem copiedItem = null;

                            try
                            {
                                item = items[i] as Outlook.MailItem;
                                if (item == null) continue;

                                if (item.Permission != Outlook.OlPermission.olUnrestricted)
                                {
                                    skippedItems++;  // Increment skipped items count
                                    continue;
                                }

                                // Check if the item is older than or equal to the cutoff date
                                if (monthsOld > 0 && item.ReceivedTime > cutoffDate)
                                {
                                    skippedItems++;  // Increment skipped items count
                                    continue;
                                }

                                if (isMoveOperation)
                                {
                                    movedItem = item.Move(destFolder);
                                }
                                else
                                {
                                    copiedItem = item.Copy();
                                    movedItem = copiedItem.Move(destFolder);
                                }
                            }
                            catch (Exception ex)
                            {
                                lastError = ex.Message;
                                errorItems++;
                                Log($"Error processing item {i}: {ex.Message}");
                            }
                            finally
                            {
                                if (item != null)
                                {
                                    Marshal.ReleaseComObject(item);
                                }
                                if (movedItem != null)
                                {
                                    Marshal.ReleaseComObject(movedItem);
                                }
                                if (copiedItem != null)
                                {
                                    Marshal.ReleaseComObject(copiedItem);
                                }

                                processedItems++;
                                progressForm.Invoke(new Action(() =>
                                {
                                    isCancelled = progressForm.UpdateProgress(processedItems, totalItems, errorItems, skippedItems, lastError, stopwatch.Elapsed);
                                }));
                            }
                        }
                    });

                    stopwatch.Stop();
                    progressForm.Close();

                    string completionMessage = $"Operation completed: {processedItems}/{totalItems} processed, {errorItems} had errors, {skippedItems} skipped";
                    Log(completionMessage);
                    Log($"Total time: {stopwatch.Elapsed}");

                    MessageBox.Show($"{completionMessage}. Last error: '{lastError}'",
                        "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                string errorMessage = $"An error occurred: {ex.Message}";
                Log(errorMessage);
                MessageBox.Show(errorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Clean up final objects
                if (items != null) Marshal.ReleaseComObject(items);
                if (destFolder != null) Marshal.ReleaseComObject(destFolder);
                if (sourceFolder != null) Marshal.ReleaseComObject(sourceFolder);
                if (pstStore != null) Marshal.ReleaseComObject(pstStore);
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
                string[] pathParts = folderPath.Split('\\');

                if (pathParts.Length > 0 && pathParts[0].Contains("@"))
                {
                    string[] temp = new string[pathParts.Length - 1];
                    Array.Copy(pathParts, 1, temp, 0, temp.Length);
                    pathParts = temp;
                }

                foreach (string part in pathParts)
                {
                    if (string.IsNullOrEmpty(part)) continue;

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
