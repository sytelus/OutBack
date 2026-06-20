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

        public PSTSelectionForm()
        {
            InitializeComponent();
            SelectedSourceFolders = new List<FolderSelection>();
            PopulateStoreComboBox();
        }

        private void PopulateStoreComboBox()
        {
            Outlook.Stores stores = null;
            try
            {
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
            Outlook.MAPIFolder currentFolder = null;
            Outlook.Store sourceStore = null;
            Outlook.MAPIFolder rootFolder = null;
            Outlook.Folders folders = null;

            checkedListFolders.Items.Clear();

            try
            {
                currentFolder = Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder;
                labelSourceFolder.Text = currentFolder.FolderPath;

                sourceStore = currentFolder.Store;
                rootFolder = sourceStore.GetRootFolder() as Outlook.MAPIFolder;
                folders = rootFolder.Folders;

                foreach (Outlook.MAPIFolder folder in folders)
                {
                    try
                    {
                        var selection = new FolderSelection(
                            folder.Name,
                            folder.FolderPath,
                            folder.EntryID,
                            folder.StoreID);

                        checkedListFolders.Items.Add(selection, true);
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
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
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

        private void PSTSelectionForm_Load(object sender, EventArgs e)
        {
            PopulateSourceFolderList();
        }
    }
}
