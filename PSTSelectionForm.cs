using System;
using System.Windows.Forms;

namespace OutBack
{
    public partial class PSTSelectionForm : Form
    {
        public string PstFilePath { get; private set; }

        public bool IsMoveOperation { get; private set; }

        public PSTSelectionForm()
        {
            InitializeComponent();
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Outlook PST Files (*.pst)|*.pst";
            dlg.Title = "Select PST File";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                txtPstFilePath.Text = dlg.FileName;
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtPstFilePath.Text))
            {
                MessageBox.Show("Please select a PST file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            PstFilePath = txtPstFilePath.Text;
            IsMoveOperation = chkMoveItems.Checked;

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
            labelSourceFolder.Text = Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder.FolderPath;
        }
    }
}
