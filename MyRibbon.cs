using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutBack
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void btnMoveCopy_Click(object sender, RibbonControlEventArgs e)
        {
            using (PSTSelectionForm pstForm = new PSTSelectionForm())
            {
                if (pstForm.ShowDialog() != DialogResult.OK)
                    return;

                string pstFilePath = pstForm.SelectedStoreName;
                bool isMoveOperation = pstForm.IsMoveOperation;
                double monthsOld = pstForm.MonthsOld;
                var sourceFolders = new List<Outlook.MAPIFolder>();
                bool passedToMover = false;

                try
                {
                    Outlook.NameSpace session = Globals.ThisAddIn.Application.Session;
                    foreach (PSTSelectionForm.FolderSelection selection in pstForm.SelectedSourceFolders)
                    {
                        sourceFolders.Add(session.GetFolderFromID(selection.EntryId, selection.StoreId));
                    }

                    ItemMover mover = new ItemMover();
                    passedToMover = true;
                    mover.Start(sourceFolders, pstFilePath, isMoveOperation, monthsOld);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Unable to start operation:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    if (!passedToMover)
                    {
                        foreach (Outlook.MAPIFolder sourceFolder in sourceFolders)
                        {
                            if (sourceFolder != null)
                                Marshal.ReleaseComObject(sourceFolder);
                        }
                    }
                }
            }
        }

        private void btnExportCalendar_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.MAPIFolder calendarFolder = null;

            try
            {
                calendarFolder = GetCalendarFolderForExport();
                using (CalendarExportForm exportForm = new CalendarExportForm(calendarFolder))
                {
                    if (exportForm.ShowDialog() != DialogResult.OK)
                        return;

                    CalendarExporter exporter = new CalendarExporter();
                    CalendarExportResult result = exporter.Export(calendarFolder, exportForm.Options);

                    MessageBox.Show(
                        $"Calendar export completed.\nExported: {result.ExportedItems}\nSkipped: {result.SkippedItems}\nErrors: {result.ErrorItems}\nFile: {result.FilePath}",
                        "Export Calendar",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to export calendar:\n{ex.Message}", "Export Calendar", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (calendarFolder != null)
                    Marshal.ReleaseComObject(calendarFolder);
            }
        }

        private void btnExportContacts_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.MAPIFolder contactFolder = null;

            try
            {
                contactFolder = GetContactFolderForExport();
                string exportFilePath = PromptForContactExportPath(contactFolder.Name);
                if (string.IsNullOrEmpty(exportFilePath))
                    return;

                ContactExporter exporter = new ContactExporter();
                ContactExportResult result = exporter.Export(contactFolder, exportFilePath);
                string status = result.Cancelled ? "cancelled" : "completed";

                MessageBox.Show(
                    $"Contact export {status}.\nExported: {result.ExportedItems}\nSkipped: {result.SkippedItems}\nErrors: {result.ErrorItems}\nFile: {result.FilePath}",
                    "Export Contacts",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unable to export contacts:\n{ex.Message}", "Export Contacts", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (contactFolder != null)
                    Marshal.ReleaseComObject(contactFolder);
            }
        }

        private Outlook.MAPIFolder GetCalendarFolderForExport()
        {
            Outlook.Explorer explorer = null;
            Outlook.MAPIFolder currentFolder = null;

            try
            {
                explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer != null)
                {
                    currentFolder = explorer.CurrentFolder;
                    if (currentFolder != null && currentFolder.DefaultItemType == Outlook.OlItemType.olAppointmentItem)
                        return currentFolder;

                    if (currentFolder != null)
                    {
                        Marshal.ReleaseComObject(currentFolder);
                        currentFolder = null;
                    }
                }

                return Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            }
            finally
            {
                if (explorer != null)
                    Marshal.ReleaseComObject(explorer);
            }
        }

        private Outlook.MAPIFolder GetContactFolderForExport()
        {
            Outlook.Explorer explorer = null;
            Outlook.MAPIFolder currentFolder = null;

            try
            {
                explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer != null)
                {
                    currentFolder = explorer.CurrentFolder;
                    if (currentFolder != null && currentFolder.DefaultItemType == Outlook.OlItemType.olContactItem)
                        return currentFolder;

                    if (currentFolder != null)
                    {
                        Marshal.ReleaseComObject(currentFolder);
                        currentFolder = null;
                    }
                }

                return Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
            }
            finally
            {
                if (explorer != null)
                    Marshal.ReleaseComObject(explorer);
            }
        }

        private static string PromptForContactExportPath(string folderName)
        {
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = "vCard files (*.vcf)|*.vcf|All files (*.*)|*.*";
                dialog.DefaultExt = "vcf";
                dialog.AddExtension = true;
                dialog.OverwritePrompt = true;
                dialog.FileName = GetDefaultContactExportFileName(folderName);
                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

                return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : string.Empty;
            }
        }

        private static string GetDefaultContactExportFileName(string folderName)
        {
            string safeFolderName = folderName;
            foreach (char invalidChar in System.IO.Path.GetInvalidFileNameChars())
            {
                safeFolderName = safeFolderName.Replace(invalidChar, '_');
            }

            return $"{safeFolderName}_{DateTime.Now:yyyyMMdd_HHmmss}.vcf";
        }
    }
}
