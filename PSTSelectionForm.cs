using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace OutBack
{
    public partial class PSTSelectionForm : Form
    {
        public sealed class FolderSelection
        {
            public string DisplayName { get; private set; }
            public string FolderPath { get; private set; }
            public string EntryId { get; private set; }
            public string StoreId { get; private set; }

            public FolderSelection(string displayName, string folderPath, string entryId, string storeId)
            {
                DisplayName = displayName;
                FolderPath = folderPath;
                EntryId = entryId;
                StoreId = storeId;
            }

            public override string ToString()
            {
                return DisplayName;
            }
        }

        public string SelectedStoreName { get; private set; }
        public bool IsMoveOperation { get; private set; }
        public double MonthsOld { get; private set; }
        public List<FolderSelection> SelectedSourceFolders { get; private set; }
        private bool isLoadingFolders;

        public PSTSelectionForm()
        {
            InitializeComponent();
            SelectedSourceFolders = new List<FolderSelection>();
            SetFolderLoadingState(true, "Loading Outlook folders...");
        }

        private void PopulateStoreComboBox()
        {
            Outlook.Stores stores = null;
            try
            {
                cboStores.Items.Clear();
                stores = Globals.ThisAddIn.Application.Session.Stores;
                foreach (Outlook.Store store in stores)
                {
                    try
                    {
                        cboStores.Items.Add(store.DisplayName);
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(store);
                    }
                }
            }
            finally
            {
                if (stores != null)
                    Marshal.ReleaseComObject(stores);
            }
        }

        private void PopulateSourceFolderList()
        {
            Outlook.Explorer explorer = null;
            Outlook.MAPIFolder currentFolder = null;
            Outlook.Store sourceStore = null;
            Outlook.MAPIFolder rootFolder = null;
            Outlook.Folders folders = null;

            checkedListFolders.Items.Clear();
            SetFolderLoadingState(true, "Loading Outlook folders...");
            Application.DoEvents();

            try
            {
                explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                currentFolder = explorer.CurrentFolder;
                labelSourceFolder.Text = currentFolder.FolderPath;

                sourceStore = currentFolder.Store;
                rootFolder = sourceStore.GetRootFolder() as Outlook.MAPIFolder;
                folders = rootFolder.Folders;
                int totalFolders = folders.Count;
                SetFolderLoadProgress(0, totalFolders);

                for (int index = 1; index <= totalFolders; index++)
                {
                    Outlook.MAPIFolder folder = null;
                    try
                    {
                        folder = folders[index] as Outlook.MAPIFolder;
                        if (folder == null)
                            continue;

                        var selection = new FolderSelection(
                            folder.Name,
                            folder.FolderPath,
                            folder.EntryID,
                            folder.StoreID);

                        checkedListFolders.Items.Add(selection, false);
                        SetFolderLoadProgress(index, totalFolders);
                    }
                    finally
                    {
                        if (folder != null)
                            Marshal.ReleaseComObject(folder);
                    }
                }
            }
            finally
            {
                if (folders != null)
                    Marshal.ReleaseComObject(folders);
                if (rootFolder != null)
                    Marshal.ReleaseComObject(rootFolder);
                if (sourceStore != null)
                    Marshal.ReleaseComObject(sourceStore);
                if (currentFolder != null)
                    Marshal.ReleaseComObject(currentFolder);
                if (explorer != null)
                    Marshal.ReleaseComObject(explorer);
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (isLoadingFolders)
            {
                MessageBox.Show("Please wait for the source folder list to finish loading.", "Loading folders", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (cboStores.SelectedItem == null)
            {
                MessageBox.Show("Please select a PST file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!double.TryParse(txtMonthsOld.Text, out double monthsOld))
            {
                MessageBox.Show("Please enter a valid number for months old.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (checkedListFolders.CheckedItems.Count == 0)
            {
                MessageBox.Show("Please select at least one source folder.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            SelectedStoreName = cboStores.SelectedItem.ToString();
            IsMoveOperation = chkMoveItems.Checked;
            MonthsOld = monthsOld;
            SelectedSourceFolders = new List<FolderSelection>();

            foreach (object item in checkedListFolders.CheckedItems)
            {
                FolderSelection selection = item as FolderSelection;
                if (selection != null)
                    SelectedSourceFolders.Add(selection);
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void PSTSelectionForm_Shown(object sender, EventArgs e)
        {
            BeginInvoke(new Action(LoadOutlookSelections));
        }

        private void LoadOutlookSelections()
        {
            try
            {
                PopulateStoreComboBox();
                PopulateSourceFolderList();
                SetFolderLoadingState(false, $"{checkedListFolders.Items.Count} source folder(s) loaded.");
            }
            catch (Exception ex)
            {
                SetFolderLoadingState(false, "Unable to load source folders.");
                MessageBox.Show($"Unable to load Outlook folders:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            SetAllSourceFolders(true);
        }

        private void btnSelectNone_Click(object sender, EventArgs e)
        {
            SetAllSourceFolders(false);
        }

        private void SetAllSourceFolders(bool isChecked)
        {
            for (int i = 0; i < checkedListFolders.Items.Count; i++)
                checkedListFolders.SetItemChecked(i, isChecked);
        }

        private void SetFolderLoadingState(bool isLoading, string statusText)
        {
            isLoadingFolders = isLoading;
            checkedListFolders.Enabled = !isLoading;
            btnSelectAll.Enabled = !isLoading && checkedListFolders.Items.Count > 0;
            btnSelectNone.Enabled = !isLoading && checkedListFolders.Items.Count > 0;
            btnOK.Enabled = !isLoading;
            btnCancel.Enabled = !isLoading;
            progressFolders.Visible = isLoading;
            progressFolders.Style = isLoading ? ProgressBarStyle.Marquee : ProgressBarStyle.Blocks;
            labelFolderLoadStatus.Text = statusText;
            Cursor = isLoading ? Cursors.WaitCursor : Cursors.Default;
        }

        private void SetFolderLoadProgress(int loadedFolders, int totalFolders)
        {
            progressFolders.Style = ProgressBarStyle.Blocks;
            progressFolders.Maximum = Math.Max(totalFolders, 1);
            progressFolders.Value = Math.Min(loadedFolders, progressFolders.Maximum);
            labelFolderLoadStatus.Text = $"Loading source folders ({loadedFolders}/{totalFolders})...";
            Application.DoEvents();
        }
    }
}
