using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutBack
{
    public sealed class ContactExportResult
    {
        public string FilePath { get; set; }
        public int TotalItems { get; set; }
        public int ExportedItems { get; set; }
        public int SkippedItems { get; set; }
        public int ErrorItems { get; set; }
        public bool Cancelled { get; set; }
    }

    public sealed class ContactExporter
    {
        private readonly string logFilePath;

        public ContactExporter()
        {
            string logDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".OutBack");
            Directory.CreateDirectory(logDirectory);
            logFilePath = Path.Combine(logDirectory, $"OutBack_ContactExport_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
        }

        public ContactExportResult Export(Outlook.MAPIFolder contactFolder, string exportFilePath)
        {
            if (contactFolder.DefaultItemType != Outlook.OlItemType.olContactItem)
                throw new InvalidOperationException("The selected folder is not a Contacts folder.");

            if (string.IsNullOrWhiteSpace(exportFilePath))
                throw new ArgumentException("An export file path is required.", "exportFilePath");

            var result = new ContactExportResult { FilePath = exportFilePath };
            Outlook.Items items = null;
            Stopwatch stopwatch = Stopwatch.StartNew();

            try
            {
                items = contactFolder.Items;
                result.TotalItems = items.Count;
                Log($"Contact export started: Source={contactFolder.FolderPath}, Items={result.TotalItems}, File={exportFilePath}");

                using (var writer = new StreamWriter(exportFilePath, false, new UTF8Encoding(false)))
                using (var progressForm = new ContactExportProgressForm(contactFolder.FolderPath, exportFilePath))
                {
                    progressForm.Show();

                    for (int i = 1; i <= result.TotalItems; i++)
                    {
                        object item = null;
                        string lastError = string.Empty;
                        bool isCancelled = false;

                        try
                        {
                            item = items[i];
                            Outlook.ContactItem contact = null;
                            if (!TryCastItem(item, out contact))
                            {
                                result.SkippedItems++;
                            }
                            else
                            {
                                writer.WriteLine(ExportContactToVCard(contact));
                                result.ExportedItems++;
                            }
                        }
                        catch (Exception ex)
                        {
                            result.ErrorItems++;
                            lastError = $"Error: {ex.GetType().Name}\nMessage: {ex.Message}";
                            Log($"Error exporting contact item {i}", ex);
                        }
                        finally
                        {
                            if (item != null) Marshal.ReleaseComObject(item);

                            isCancelled = progressForm.UpdateProgress(
                                i,
                                result.TotalItems,
                                result.ExportedItems,
                                result.SkippedItems,
                                result.ErrorItems,
                                lastError,
                                stopwatch.Elapsed);
                        }

                        if (isCancelled)
                        {
                            result.Cancelled = true;
                            break;
                        }
                    }

                    progressForm.Close();
                }
            }
            finally
            {
                if (items != null) Marshal.ReleaseComObject(items);
            }

            stopwatch.Stop();
            Log($"Contact export completed: Exported={result.ExportedItems}, Skipped={result.SkippedItems}, Errors={result.ErrorItems}, Time={stopwatch.Elapsed}");
            return result;
        }

        private static string ExportContactToVCard(Outlook.ContactItem contact)
        {
            string tempPath = Path.Combine(Path.GetTempPath(), $"OutBack_{Guid.NewGuid():N}.vcf");

            try
            {
                contact.SaveAs(tempPath, Outlook.OlSaveAsType.olVCard);
                return File.ReadAllText(tempPath, Encoding.Default).TrimEnd('\r', '\n');
            }
            finally
            {
                try
                {
                    if (File.Exists(tempPath))
                        File.Delete(tempPath);
                }
                catch
                {
                }
            }
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

        private void Log(string message, Exception ex = null)
        {
            string logEntry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
            if (ex != null)
            {
                logEntry += $"\nException: {ex.GetType().FullName}" +
                           $"\nMessage: {ex.Message}" +
                           $"\nStack Trace:\n{ex.StackTrace}";
            }

            File.AppendAllText(logFilePath, logEntry + Environment.NewLine + Environment.NewLine);
        }

        private sealed class ContactExportProgressForm : Form
        {
            private readonly ProgressBar progressBar;
            private readonly Label labelStatus;
            private readonly Label labelSource;
            private readonly Label labelDestination;
            private readonly Label labelElapsed;
            private readonly TextBox txtLastError;
            private bool isCancelled;

            public ContactExportProgressForm(string sourceFolder, string exportFilePath)
            {
                Text = "Export Contacts";
                ClientSize = new System.Drawing.Size(580, 260);
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                StartPosition = FormStartPosition.CenterScreen;

                labelSource = new Label
                {
                    AutoEllipsis = true,
                    Location = new System.Drawing.Point(16, 16),
                    Size = new System.Drawing.Size(548, 24),
                    Text = $"Source: {sourceFolder}"
                };

                labelDestination = new Label
                {
                    AutoEllipsis = true,
                    Location = new System.Drawing.Point(16, 44),
                    Size = new System.Drawing.Size(548, 24),
                    Text = $"File: {exportFilePath}"
                };

                progressBar = new ProgressBar
                {
                    Location = new System.Drawing.Point(16, 80),
                    Size = new System.Drawing.Size(548, 24)
                };

                labelStatus = new Label
                {
                    AutoSize = true,
                    Location = new System.Drawing.Point(16, 116),
                    Text = "Starting export..."
                };

                labelElapsed = new Label
                {
                    AutoSize = true,
                    Location = new System.Drawing.Point(16, 144),
                    Text = "Elapsed: 00:00:00"
                };

                txtLastError = new TextBox
                {
                    Location = new System.Drawing.Point(16, 176),
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Vertical,
                    Size = new System.Drawing.Size(428, 64)
                };

                Button btnCancel = new Button
                {
                    Location = new System.Drawing.Point(464, 201),
                    Size = new System.Drawing.Size(100, 39),
                    Text = "Cancel",
                    UseVisualStyleBackColor = true
                };
                btnCancel.Click += (sender, args) => isCancelled = true;

                Controls.Add(labelSource);
                Controls.Add(labelDestination);
                Controls.Add(progressBar);
                Controls.Add(labelStatus);
                Controls.Add(labelElapsed);
                Controls.Add(txtLastError);
                Controls.Add(btnCancel);
            }

            public bool UpdateProgress(int processedItems, int totalItems, int exportedItems, int skippedItems, int errorItems, string lastError, TimeSpan elapsedTime)
            {
                progressBar.Value = totalItems <= 0 ? 100 : Math.Min((int)((double)processedItems / totalItems * 100), 100);
                labelStatus.Text = $"Processed {processedItems} of {totalItems}. Exported: {exportedItems}. Skipped: {skippedItems}. Errors: {errorItems}.";
                labelElapsed.Text = $"Elapsed: {elapsedTime:hh\\:mm\\:ss}";

                if (!string.IsNullOrEmpty(lastError))
                    txtLastError.Text = lastError;

                Application.DoEvents();
                return isCancelled;
            }
        }
    }
}
