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
    }
}
