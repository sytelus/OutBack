using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutBack
{
    public class ItemMover
    {
        private const int BATCH_SIZE = 20; // Process items in smaller batches

        public async void Start(Outlook.MAPIFolder sourceFolder, string pstFilePath, bool isMoveOperation)
        {
            try
            {
                // Open or add the PST file
                Outlook.Stores stores = Globals.ThisAddIn.Application.Session.Stores;
                Outlook.Store pstStore = null;

                foreach (Outlook.Store store in stores)
                {
                    try
                    {
                        if (store.FilePath == pstFilePath)
                        {
                            pstStore = store;
                            break;
                        }
                    }
                    finally
                    {
                        if (store != pstStore) Marshal.ReleaseComObject(store);
                    }
                }
                Marshal.ReleaseComObject(stores);

                if (pstStore == null)
                {
                    throw new Exception("PST file not found");
                }

                // Get or create the destination folder
                Outlook.MAPIFolder destFolder = GetOrCreateFolder(pstStore, sourceFolder.FolderPath);
                Marshal.ReleaseComObject(pstStore);

                // Get all items in the source folder
                Outlook.Items items = sourceFolder.Items;
                int totalItems = items.Count;
                int processedItems = 0;
                int errorItems = 0;

                using (ProgressForm progressForm = new ProgressForm())
                {
                    progressForm.Show();
                    Stopwatch stopwatch = Stopwatch.StartNew();
                    bool isCancelled = false;
                    string lastError = "";

                    await Task.Run(() =>
                    {
                        for (int i = totalItems; i > 0; i--)
                        {
                            if (isCancelled) break;

                            if ((i % BATCH_SIZE == 0) || (i == totalItems))
                            {
                                // Force garbage collection periodically
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                            }

                            Outlook.MailItem item = null;
                            try
                            {
                                item = items[i] as Outlook.MailItem;
                                if (item == null) continue;

                                if (isMoveOperation)
                                {
                                    Outlook.MailItem movedItem = item.Move(destFolder);
                                    if (movedItem != null) Marshal.ReleaseComObject(movedItem);
                                }
                                else
                                {
                                    Outlook.MailItem copiedItem = item.Copy();
                                    Outlook.MailItem movedItem = copiedItem.Move(destFolder);
                                    if (copiedItem != null) Marshal.ReleaseComObject(copiedItem);
                                    if (movedItem != null) Marshal.ReleaseComObject(movedItem);
                                }
                            }
                            catch (Exception ex)
                            {
                                lastError = ex.Message;
                                errorItems++;
                                continue;
                            }
                            finally
                            {
                                if (item != null)
                                {
                                    item.Close(Outlook.OlInspectorClose.olDiscard);
                                    Marshal.ReleaseComObject(item);
                                }

                                processedItems++;
                                progressForm.Invoke(new Action(() =>
                                {
                                    isCancelled = progressForm.UpdateProgress(processedItems, totalItems, errorItems, lastError, stopwatch.Elapsed);
                                }));
                            }
                        }
                    });

                    stopwatch.Stop();
                    progressForm.Close();

                    // Clean up final objects
                    Marshal.ReleaseComObject(items);
                    Marshal.ReleaseComObject(destFolder);
                    Marshal.ReleaseComObject(sourceFolder);

                    MessageBox.Show($"Operation completed: {processedItems}/{totalItems} processed, {errorItems} had errors such as '{lastError}'",
                        "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Outlook.MAPIFolder GetOrCreateFolder(Outlook.Store store, string folderPath)
        {
            Outlook.Folder rootFolder = store.GetRootFolder() as Outlook.Folder;
            string[] pathParts = folderPath.Split('\\');
            Outlook.Folder currentFolder = rootFolder;

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
                    if (currentFolder != rootFolder) Marshal.ReleaseComObject(currentFolder);
                    currentFolder = subFolder;
                }
                catch
                {
                    subFolder = currentFolder.Folders.Add(part, Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
                    if (currentFolder != rootFolder) Marshal.ReleaseComObject(currentFolder);
                    currentFolder = subFolder;
                }
            }

            Marshal.ReleaseComObject(rootFolder);
            return currentFolder;
        }
    }
}