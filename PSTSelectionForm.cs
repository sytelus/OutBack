using System;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace OutBack
{
    public partial class PSTSelectionForm : Form
    {
        public string SelectedStoreName { get; private set; }
        public bool IsMoveOperation { get; private set; }
        public double MonthsOld { get; private set; }

        public PSTSelectionForm()
        {
            InitializeComponent();
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

            SelectedStoreName = cboStores.SelectedItem.ToString();
            IsMoveOperation = chkMoveItems.Checked;
            MonthsOld = monthsOld;

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
            Outlook.MAPIFolder folder = null;
            try
            {
                folder = Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder;
                labelSourceFolder.Text = folder.FolderPath;
            }
            finally
            {
                if (folder != null)
                    Marshal.ReleaseComObject(folder);
            }
        }
    }
}
