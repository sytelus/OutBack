namespace OutBack
{
    partial class CalendarExportForm
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.Label lblCalendarFolder;
        private System.Windows.Forms.Label labelCalendarFolder;
        private System.Windows.Forms.Label lblExportFile;
        private System.Windows.Forms.TextBox txtExportFile;
        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.Label lblCategories;
        private System.Windows.Forms.CheckedListBox checkedListCategories;
        private System.Windows.Forms.Button btnCheckAllCategories;
        private System.Windows.Forms.Button btnClearCategories;
        private System.Windows.Forms.CheckBox chkAppointmentsOnly;
        private System.Windows.Forms.CheckBox chkMeetingsOnly;
        private System.Windows.Forms.CheckBox chkOrganizedByCurrentUserOnly;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.lblCalendarFolder = new System.Windows.Forms.Label();
            this.labelCalendarFolder = new System.Windows.Forms.Label();
            this.lblExportFile = new System.Windows.Forms.Label();
            this.txtExportFile = new System.Windows.Forms.TextBox();
            this.btnBrowse = new System.Windows.Forms.Button();
            this.lblCategories = new System.Windows.Forms.Label();
            this.checkedListCategories = new System.Windows.Forms.CheckedListBox();
            this.btnCheckAllCategories = new System.Windows.Forms.Button();
            this.btnClearCategories = new System.Windows.Forms.Button();
            this.chkAppointmentsOnly = new System.Windows.Forms.CheckBox();
            this.chkMeetingsOnly = new System.Windows.Forms.CheckBox();
            this.chkOrganizedByCurrentUserOnly = new System.Windows.Forms.CheckBox();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblCalendarFolder
            // 
            this.lblCalendarFolder.AutoSize = true;
            this.lblCalendarFolder.Location = new System.Drawing.Point(12, 17);
            this.lblCalendarFolder.Name = "lblCalendarFolder";
            this.lblCalendarFolder.Size = new System.Drawing.Size(124, 20);
            this.lblCalendarFolder.TabIndex = 0;
            this.lblCalendarFolder.Text = "Calendar folder:";
            // 
            // labelCalendarFolder
            // 
            this.labelCalendarFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelCalendarFolder.AutoEllipsis = true;
            this.labelCalendarFolder.Location = new System.Drawing.Point(150, 17);
            this.labelCalendarFolder.Name = "labelCalendarFolder";
            this.labelCalendarFolder.Size = new System.Drawing.Size(452, 24);
            this.labelCalendarFolder.TabIndex = 1;
            this.labelCalendarFolder.Text = "Calendar";
            // 
            // lblExportFile
            // 
            this.lblExportFile.AutoSize = true;
            this.lblExportFile.Location = new System.Drawing.Point(12, 58);
            this.lblExportFile.Name = "lblExportFile";
            this.lblExportFile.Size = new System.Drawing.Size(82, 20);
            this.lblExportFile.TabIndex = 2;
            this.lblExportFile.Text = "Export file:";
            // 
            // txtExportFile
            // 
            this.txtExportFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtExportFile.Location = new System.Drawing.Point(150, 55);
            this.txtExportFile.Name = "txtExportFile";
            this.txtExportFile.Size = new System.Drawing.Size(351, 26);
            this.txtExportFile.TabIndex = 3;
            // 
            // btnBrowse
            // 
            this.btnBrowse.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnBrowse.Location = new System.Drawing.Point(507, 53);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(95, 31);
            this.btnBrowse.TabIndex = 4;
            this.btnBrowse.Text = "Browse...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // lblCategories
            // 
            this.lblCategories.AutoSize = true;
            this.lblCategories.Location = new System.Drawing.Point(12, 104);
            this.lblCategories.Name = "lblCategories";
            this.lblCategories.Size = new System.Drawing.Size(91, 20);
            this.lblCategories.TabIndex = 5;
            this.lblCategories.Text = "Categories:";
            // 
            // checkedListCategories
            // 
            this.checkedListCategories.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.checkedListCategories.CheckOnClick = true;
            this.checkedListCategories.FormattingEnabled = true;
            this.checkedListCategories.Location = new System.Drawing.Point(16, 130);
            this.checkedListCategories.Name = "checkedListCategories";
            this.checkedListCategories.Size = new System.Drawing.Size(586, 172);
            this.checkedListCategories.TabIndex = 6;
            // 
            // btnCheckAllCategories
            // 
            this.btnCheckAllCategories.Location = new System.Drawing.Point(16, 312);
            this.btnCheckAllCategories.Name = "btnCheckAllCategories";
            this.btnCheckAllCategories.Size = new System.Drawing.Size(95, 31);
            this.btnCheckAllCategories.TabIndex = 7;
            this.btnCheckAllCategories.Text = "Check All";
            this.btnCheckAllCategories.UseVisualStyleBackColor = true;
            this.btnCheckAllCategories.Click += new System.EventHandler(this.btnCheckAllCategories_Click);
            // 
            // btnClearCategories
            // 
            this.btnClearCategories.Location = new System.Drawing.Point(117, 312);
            this.btnClearCategories.Name = "btnClearCategories";
            this.btnClearCategories.Size = new System.Drawing.Size(95, 31);
            this.btnClearCategories.TabIndex = 8;
            this.btnClearCategories.Text = "Clear";
            this.btnClearCategories.UseVisualStyleBackColor = true;
            this.btnClearCategories.Click += new System.EventHandler(this.btnClearCategories_Click);
            // 
            // chkAppointmentsOnly
            // 
            this.chkAppointmentsOnly.AutoSize = true;
            this.chkAppointmentsOnly.Location = new System.Drawing.Point(16, 366);
            this.chkAppointmentsOnly.Name = "chkAppointmentsOnly";
            this.chkAppointmentsOnly.Size = new System.Drawing.Size(170, 24);
            this.chkAppointmentsOnly.TabIndex = 9;
            this.chkAppointmentsOnly.Text = "Appointments only";
            this.chkAppointmentsOnly.UseVisualStyleBackColor = true;
            this.chkAppointmentsOnly.CheckedChanged += new System.EventHandler(this.chkAppointmentsOnly_CheckedChanged);
            // 
            // chkMeetingsOnly
            // 
            this.chkMeetingsOnly.AutoSize = true;
            this.chkMeetingsOnly.Location = new System.Drawing.Point(218, 366);
            this.chkMeetingsOnly.Name = "chkMeetingsOnly";
            this.chkMeetingsOnly.Size = new System.Drawing.Size(131, 24);
            this.chkMeetingsOnly.TabIndex = 10;
            this.chkMeetingsOnly.Text = "Meetings only";
            this.chkMeetingsOnly.UseVisualStyleBackColor = true;
            this.chkMeetingsOnly.CheckedChanged += new System.EventHandler(this.chkMeetingsOnly_CheckedChanged);
            // 
            // chkOrganizedByCurrentUserOnly
            // 
            this.chkOrganizedByCurrentUserOnly.AutoSize = true;
            this.chkOrganizedByCurrentUserOnly.Location = new System.Drawing.Point(16, 404);
            this.chkOrganizedByCurrentUserOnly.Name = "chkOrganizedByCurrentUserOnly";
            this.chkOrganizedByCurrentUserOnly.Size = new System.Drawing.Size(247, 24);
            this.chkOrganizedByCurrentUserOnly.TabIndex = 11;
            this.chkOrganizedByCurrentUserOnly.Text = "Organized by current user only";
            this.chkOrganizedByCurrentUserOnly.UseVisualStyleBackColor = true;
            // 
            // btnOK
            // 
            this.btnOK.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOK.Location = new System.Drawing.Point(400, 456);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(98, 39);
            this.btnOK.TabIndex = 12;
            this.btnOK.Text = "Export";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.Location = new System.Drawing.Point(504, 456);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(98, 39);
            this.btnCancel.TabIndex = 13;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // CalendarExportForm
            // 
            this.ClientSize = new System.Drawing.Size(614, 512);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.chkOrganizedByCurrentUserOnly);
            this.Controls.Add(this.chkMeetingsOnly);
            this.Controls.Add(this.chkAppointmentsOnly);
            this.Controls.Add(this.btnClearCategories);
            this.Controls.Add(this.btnCheckAllCategories);
            this.Controls.Add(this.checkedListCategories);
            this.Controls.Add(this.lblCategories);
            this.Controls.Add(this.btnBrowse);
            this.Controls.Add(this.txtExportFile);
            this.Controls.Add(this.lblExportFile);
            this.Controls.Add(this.labelCalendarFolder);
            this.Controls.Add(this.lblCalendarFolder);
            this.Name = "CalendarExportForm";
            this.Text = "Export Calendar";
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
