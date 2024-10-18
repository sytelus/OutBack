namespace OutBack
{
    partial class PSTSelectionForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Label lblPstFile;
        private System.Windows.Forms.TextBox txtPstFilePath;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.CheckBox chkMoveItems;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;

        private void InitializeComponent()
        {
            this.lblPstFile = new System.Windows.Forms.Label();
            this.txtPstFilePath = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.chkMoveItems = new System.Windows.Forms.CheckBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            //
            // lblPstFile
            //
            this.lblPstFile.AutoSize = true;
            this.lblPstFile.Location = new System.Drawing.Point(12, 15);
            this.lblPstFile.Name = "lblPstFile";
            this.lblPstFile.Size = new System.Drawing.Size(50, 13);
            this.lblPstFile.TabIndex = 0;
            this.lblPstFile.Text = "PST File:";
            //
            // txtPstFilePath
            //
            this.txtPstFilePath.Location = new System.Drawing.Point(68, 12);
            this.txtPstFilePath.Name = "txtPstFilePath";
            this.txtPstFilePath.Size = new System.Drawing.Size(250, 20);
            this.txtPstFilePath.TabIndex = 1;
            //
            // btnBrowse
            //
            this.btnBrowse.Location = new System.Drawing.Point(324, 10);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(75, 23);
            this.btnBrowse.TabIndex = 2;
            this.btnBrowse.Text = "Browse...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            //
            // chkMoveItems
            //
            this.chkMoveItems.AutoSize = true;
            this.chkMoveItems.Location = new System.Drawing.Point(68, 48);
            this.chkMoveItems.Name = "chkMoveItems";
            this.chkMoveItems.Size = new System.Drawing.Size(80, 17);
            this.chkMoveItems.TabIndex = 3;
            this.chkMoveItems.Text = "Move Items";
            this.chkMoveItems.UseVisualStyleBackColor = true;
            //
            // btnOK
            //
            this.btnOK.Location = new System.Drawing.Point(243, 82);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 4;
            this.btnOK.Text = "Start";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            //
            // btnCancel
            //
            this.btnCancel.Location = new System.Drawing.Point(324, 82);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 5;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            //
            // PSTSelectionForm
            //
            this.ClientSize = new System.Drawing.Size(411, 117);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.chkMoveItems);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.txtPstFilePath);
            this.Controls.Add(this.lblPstFile);
            this.Name = "PSTSelectionForm";
            this.Text = "Select PST File";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
