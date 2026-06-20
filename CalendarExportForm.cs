using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutBack
{
    public partial class CalendarExportForm : Form
    {
        private const string NoCategoryLabel = "(No category)";

        public CalendarExportOptions Options { get; private set; }

        public CalendarExportForm(Outlook.MAPIFolder calendarFolder)
        {
            InitializeComponent();
            Options = new CalendarExportOptions();

            labelCalendarFolder.Text = calendarFolder.FolderPath;
            txtExportFile.Text = GetDefaultExportPath(calendarFolder.Name);
            PopulateCategories();
        }

        private void PopulateCategories()
        {
            var categories = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
            Outlook.Categories outlookCategories = null;

            try
            {
                outlookCategories = Globals.ThisAddIn.Application.Session.Categories;
                foreach (Outlook.Category category in outlookCategories)
                {
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(category.Name))
                            categories.Add(category.Name);
                    }
                    finally
                    {
                        if (category != null)
                            Marshal.ReleaseComObject(category);
                    }
                }
            }
            finally
            {
                if (outlookCategories != null) Marshal.ReleaseComObject(outlookCategories);
            }

            checkedListCategories.Items.Clear();
            foreach (string category in categories)
            {
                checkedListCategories.Items.Add(category, true);
            }

            checkedListCategories.Items.Add(NoCategoryLabel, true);
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Filter = "iCalendar files (*.ics)|*.ics|All files (*.*)|*.*";
                dialog.DefaultExt = "ics";
                dialog.AddExtension = true;
                dialog.OverwritePrompt = true;

                string currentPath = txtExportFile.Text.Trim();
                if (!string.IsNullOrEmpty(currentPath))
                {
                    string directory = Path.GetDirectoryName(currentPath);
                    if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory))
                        dialog.InitialDirectory = directory;

                    dialog.FileName = Path.GetFileName(currentPath);
                }

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtExportFile.Text = dialog.FileName;
                }
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            string exportPath = txtExportFile.Text.Trim();
            if (string.IsNullOrEmpty(exportPath))
            {
                MessageBox.Show("Please select an export file.", "Export Calendar", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrEmpty(Path.GetDirectoryName(exportPath)))
                exportPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), exportPath);

            if (!Path.HasExtension(exportPath))
                exportPath = Path.ChangeExtension(exportPath, ".ics");

            string directory = Path.GetDirectoryName(exportPath);
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
            {
                MessageBox.Show("The selected export folder does not exist.", "Export Calendar", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var selectedCategories = new List<string>();
            bool includeUncategorized = false;

            foreach (object item in checkedListCategories.CheckedItems)
            {
                string category = item.ToString();
                if (string.Equals(category, NoCategoryLabel, StringComparison.OrdinalIgnoreCase))
                    includeUncategorized = true;
                else
                    selectedCategories.Add(category);
            }

            Options = new CalendarExportOptions
            {
                ExportFilePath = exportPath,
                FilterByCategories = checkedListCategories.Items.Count > 0 &&
                    checkedListCategories.CheckedItems.Count < checkedListCategories.Items.Count,
                IncludeUncategorized = includeUncategorized,
                AppointmentsOnly = chkAppointmentsOnly.Checked,
                MeetingsOnly = chkMeetingsOnly.Checked,
                OrganizedByCurrentUserOnly = chkOrganizedByCurrentUserOnly.Checked
            };
            Options.Categories.AddRange(selectedCategories);

            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void btnCheckAllCategories_Click(object sender, EventArgs e)
        {
            SetAllCategories(true);
        }

        private void btnClearCategories_Click(object sender, EventArgs e)
        {
            SetAllCategories(false);
        }

        private void chkAppointmentsOnly_CheckedChanged(object sender, EventArgs e)
        {
            if (chkAppointmentsOnly.Checked)
                chkMeetingsOnly.Checked = false;
        }

        private void chkMeetingsOnly_CheckedChanged(object sender, EventArgs e)
        {
            if (chkMeetingsOnly.Checked)
                chkAppointmentsOnly.Checked = false;
        }

        private void SetAllCategories(bool isChecked)
        {
            for (int i = 0; i < checkedListCategories.Items.Count; i++)
            {
                checkedListCategories.SetItemChecked(i, isChecked);
            }
        }

        private static string GetDefaultExportPath(string folderName)
        {
            string safeFolderName = folderName;
            foreach (char invalidChar in Path.GetInvalidFileNameChars())
            {
                safeFolderName = safeFolderName.Replace(invalidChar, '_');
            }

            string fileName = $"{safeFolderName}_{DateTime.Now:yyyyMMdd_HHmmss}.ics";
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), fileName);
        }

    }
}
