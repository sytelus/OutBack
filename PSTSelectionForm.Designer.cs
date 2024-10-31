namespace OutBack
{
    partial class PSTSelectionForm
    {
        private System.Windows.Forms.Label lblPstFile;
        private System.Windows.Forms.ComboBox cboStores;
        private System.Windows.Forms.CheckBox chkMoveItems;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label lblMonthsOld;
        private System.Windows.Forms.TextBox txtMonthsOld;

        private void InitializeComponent()
        {
            this.lblPstFile = new System.Windows.Forms.Label();
            this.cboStores = new System.Windows.Forms.ComboBox();
            this.chkMoveItems = new System.Windows.Forms.CheckBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.labelSourceFolder = new System.Windows.Forms.Label();
            this.lblMonthsOld = new System.Windows.Forms.Label();
            this.txtMonthsOld = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // lblPstFile
            // 
            this.lblPstFile.AutoSize = true;
            this.lblPstFile.Location = new System.Drawing.Point(12, 15);
            this.lblPstFile.Name = "lblPstFile";
            this.lblPstFile.Size = new System.Drawing.Size(72, 20);
            this.lblPstFile.TabIndex = 0;
            this.lblPstFile.Text = "PST File:";
            // 
            // cboStores
            // 
            this.cboStores.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboStores.FormattingEnabled = true;
            this.cboStores.Location = new System.Drawing.Point(100, 12);
            this.cboStores.Name = "cboStores";
            this.cboStores.Size = new System.Drawing.Size(410, 28);
            this.cboStores.TabIndex = 1;
            // 
            // chkMoveItems
            // 
            this.chkMoveItems.AutoSize = true;
            this.chkMoveItems.Checked = true;
            this.chkMoveItems.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkMoveItems.Location = new System.Drawing.Point(68, 48);
            this.chkMoveItems.Name = "chkMoveItems";
            this.chkMoveItems.Size = new System.Drawing.Size(117, 24);
            this.chkMoveItems.TabIndex = 3;
            this.chkMoveItems.Text = "Move Items";
            this.chkMoveItems.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(306, 169);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(98, 39);
            this.btnOK.TabIndex = 4;
            this.btnOK.Text = "Start";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(410, 169);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(100, 39);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 119);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(58, 20);
            this.label1.TabIndex = 6;
            this.label1.Text = "Folder:";
            // 
            // labelSourceFolder
            // 
            this.labelSourceFolder.AutoSize = true;
            this.labelSourceFolder.Location = new System.Drawing.Point(96, 119);
            this.labelSourceFolder.Name = "labelSourceFolder";
            this.labelSourceFolder.Size = new System.Drawing.Size(51, 20);
            this.labelSourceFolder.TabIndex = 7;
            this.labelSourceFolder.Text = "label2";
            // 
            // lblMonthsOld
            // 
            this.lblMonthsOld.AutoSize = true;
            this.lblMonthsOld.Location = new System.Drawing.Point(12, 80);
            this.lblMonthsOld.Name = "lblMonthsOld";
            this.lblMonthsOld.Size = new System.Drawing.Size(126, 20);
            this.lblMonthsOld.TabIndex = 8;
            this.lblMonthsOld.Text = "Only months old:";
            // 
            // txtMonthsOld
            // 
            this.txtMonthsOld.Location = new System.Drawing.Point(158, 77);
            this.txtMonthsOld.Name = "txtMonthsOld";
            this.txtMonthsOld.Size = new System.Drawing.Size(100, 26);
            this.txtMonthsOld.TabIndex = 9;
            this.txtMonthsOld.Text = "0";
            // 
            // PSTSelectionForm
            // 
            this.ClientSize = new System.Drawing.Size(522, 220);
            this.Controls.Add(this.txtMonthsOld);
            this.Controls.Add(this.lblMonthsOld);
            this.Controls.Add(this.labelSourceFolder);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.chkMoveItems);
            this.Controls.Add(this.cboStores);
            this.Controls.Add(this.lblPstFile);
            this.Name = "PSTSelectionForm";
            this.Text = "Select PST File";
            this.Load += new System.EventHandler(this.PSTSelectionForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label labelSourceFolder;
    }
}
