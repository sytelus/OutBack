using System;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutBack
{
    public class ItemMover
    {
        public async void Start(Outlook.MAPIFolder sourceFolder, string pstFilePath, bool isMoveOperation)
        {
            try
            {
                // Open or add the PST file
                Outlook.Stores stores = Globals.ThisAddIn.Application.Session.Stores;
                Outlook.Store pstStore = null;

                foreach (Outlook.Store store in stores)
                {
                    if (store.FilePath == pstFilePath)
                    {
                        pstStore = store;
                        break;
                    }
                }

                if (pstStore == null)
                {
                    throw new Exception("PST file not found");
                }

                // Get or create the destination folder
                Outlook.MAPIFolder destFolder = GetOrCreateFolder(pstStore, sourceFolder.FolderPath);

                // Get all items in the source folder
                Outlook.Items items = sourceFolder.Items;
                int totalItems = items.Count;
                int processedItems = 0;
                int errorItems = 0;

                ProgressForm progressForm = new ProgressForm();
                progressForm.Show();

                Stopwatch stopwatch = Stopwatch.StartNew();
                bool isCancelled = false;
                string lastError = "";

                await Task.Run(() =>
                {
                    foreach (Outlook.MailItem item in items)
                    {
                        if (isCancelled) break;
                        try
                        {
                            //// Ensure item is fully downloaded
                            //if (item is Outlook._MailItem mailItem)
                            //{
                            //    if ((mailItem.Conflicts != null && mailItem.Conflicts.Count > 0) || mailItem.IsMarkedAsTask)
                            //    {
                            //        // Skip problematic items
                            //        continue;
                            //    }
                            //}

                            if (isMoveOperation)
                            {
                                item.Move(destFolder);
                            }
                            else
                            {
                                item.Copy().Move(destFolder);
                            }
                        }
                        catch(Exception ex)
                        {
                            lastError = ex.Message;
                            // Ignore items that cannot be copied or moved
                            errorItems++;
                            continue;
                        }
                        finally
                        {
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
                MessageBox.Show($"Operation completed: {processedItems}/{totalItems} processed, {errorItems} had errors such as '{lastError}'", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                    // Folder does not exist, create it
                    subFolder = currentFolder.Folders.Add(part, Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
                }
                currentFolder = subFolder;
            }
            return currentFolder;
        }
    }
}
