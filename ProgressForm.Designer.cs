namespace OutBack
{
    partial class ProgressForm
    {
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Label lblTimeElapsed;
        private System.Windows.Forms.Label lblTimeRemaining;

        private void InitializeComponent()
        {
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.lblStatus = new System.Windows.Forms.Label();
            this.lblTimeElapsed = new System.Windows.Forms.Label();
            this.lblTimeRemaining = new System.Windows.Forms.Label();
            this.labelErrorItems = new System.Windows.Forms.Label();
            this.labelLastError = new System.Windows.Forms.Label();
            this.txtLastError = new System.Windows.Forms.TextBox();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.labelSource = new System.Windows.Forms.Label();
            this.labelDestination = new System.Windows.Forms.Label();
            this.labelSkippedItems = new System.Windows.Forms.Label();
            this.labelRetries = new System.Windows.Forms.Label();
            this.labelSkippedCast = new System.Windows.Forms.Label();
            this.labelSkippedPermissions = new System.Windows.Forms.Label();
            this.labelSkippedDate = new System.Windows.Forms.Label();
            this.labelSkippedExisting = new System.Windows.Forms.Label();
            this.labelReplacedExisting = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(15, 25);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(360, 23);
            this.progressBar.TabIndex = 0;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(12, 61);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(143, 20);
            this.lblStatus.TabIndex = 1;
            this.lblStatus.Text = "Processed 0 items.";
            // 
            // lblTimeElapsed
            // 
            this.lblTimeElapsed.AutoSize = true;
            this.lblTimeElapsed.Location = new System.Drawing.Point(12, 84);
            this.lblTimeElapsed.Name = "lblTimeElapsed";
            this.lblTimeElapsed.Size = new System.Drawing.Size(113, 20);
            this.lblTimeElapsed.TabIndex = 2;
            this.lblTimeElapsed.Text = "Time Elapsed: ";
            // 
            // lblTimeRemaining
            // 
            this.lblTimeRemaining.AutoSize = true;
            this.lblTimeRemaining.Location = new System.Drawing.Point(12, 107);
            this.lblTimeRemaining.Name = "lblTimeRemaining";
            this.lblTimeRemaining.Size = new System.Drawing.Size(165, 20);
            this.lblTimeRemaining.TabIndex = 3;
            this.lblTimeRemaining.Text = "Estimated Remaining:";
            // 
            // labelErrorItems
            // 
            this.labelErrorItems.AutoSize = true;
            this.labelErrorItems.Location = new System.Drawing.Point(12, 262);
            this.labelErrorItems.Name = "labelErrorItems";
            this.labelErrorItems.Size = new System.Drawing.Size(115, 20);
            this.labelErrorItems.TabIndex = 4;
            this.labelErrorItems.Text = "Errors: 0 items.";
            // 
            // labelLastError
            // 
            this.labelLastError.AutoSize = true;
            this.labelLastError.Location = new System.Drawing.Point(11, 341);
            this.labelLastError.Name = "labelLastError";
            this.labelLastError.Size = new System.Drawing.Size(76, 20);
            this.labelLastError.TabIndex = 5;
            this.labelLastError.Text = "Last Error";
            //
            // txtLastError
            //
            this.txtLastError.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtLastError.Location = new System.Drawing.Point(15, 364);
            this.txtLastError.Multiline = true;
            this.txtLastError.Name = "txtLastError";
            this.txtLastError.ReadOnly = true;
            this.txtLastError.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtLastError.Size = new System.Drawing.Size(850, 405);
            this.txtLastError.TabIndex = 14;
            this.txtLastError.WordWrap = false;
            // 
            // buttonCancel
            // 
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(554, 25);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(152, 46);
            this.buttonCancel.TabIndex = 6;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.button1_Click);
            // 
            // labelSource
            // 
            this.labelSource.AutoSize = true;
            this.labelSource.Location = new System.Drawing.Point(11, 146);
            this.labelSource.Name = "labelSource";
            this.labelSource.Size = new System.Drawing.Size(57, 20);
            this.labelSource.TabIndex = 7;
            this.labelSource.Text = "source";
            // 
            // labelDestination
            // 
            this.labelDestination.AutoSize = true;
            this.labelDestination.Location = new System.Drawing.Point(12, 176);
            this.labelDestination.Name = "labelDestination";
            this.labelDestination.Size = new System.Drawing.Size(87, 20);
            this.labelDestination.TabIndex = 8;
            this.labelDestination.Text = "destination";
            // 
            // labelSkippedItems
            // 
            this.labelSkippedItems.AutoSize = true;
            this.labelSkippedItems.Location = new System.Drawing.Point(12, 221);
            this.labelSkippedItems.Name = "labelSkippedItems";
            this.labelSkippedItems.Size = new System.Drawing.Size(169, 20);
            this.labelSkippedItems.TabIndex = 9;
            this.labelSkippedItems.Text = "Skipped Total: 0 items.";
            // 
            // labelRetries
            // 
            this.labelRetries.AutoSize = true;
            this.labelRetries.Location = new System.Drawing.Point(12, 297);
            this.labelRetries.Name = "labelRetries";
            this.labelRetries.Size = new System.Drawing.Size(60, 20);
            this.labelRetries.TabIndex = 10;
            this.labelRetries.Text = "Retries";
            // 
            // labelSkippedCast
            // 
            this.labelSkippedCast.AutoSize = true;
            this.labelSkippedCast.Location = new System.Drawing.Point(278, 221);
            this.labelSkippedCast.Name = "labelSkippedCast";
            this.labelSkippedCast.Size = new System.Drawing.Size(167, 20);
            this.labelSkippedCast.TabIndex = 11;
            this.labelSkippedCast.Text = "Skipped Cast: 0 items.";
            // 
            // labelSkippedPermissions
            // 
            this.labelSkippedPermissions.AutoSize = true;
            this.labelSkippedPermissions.Location = new System.Drawing.Point(473, 221);
            this.labelSkippedPermissions.Name = "labelSkippedPermissions";
            this.labelSkippedPermissions.Size = new System.Drawing.Size(179, 20);
            this.labelSkippedPermissions.TabIndex = 12;
            this.labelSkippedPermissions.Text = "Skipped Perms: 0 items.";
            // 
            // labelSkippedDate
            // 
            this.labelSkippedDate.AutoSize = true;
            this.labelSkippedDate.Location = new System.Drawing.Point(683, 221);
            this.labelSkippedDate.Name = "labelSkippedDate";
            this.labelSkippedDate.Size = new System.Drawing.Size(169, 20);
            this.labelSkippedDate.TabIndex = 13;
            this.labelSkippedDate.Text = "Skipped Date: 0 items.";
            //
            // labelSkippedExisting
            //
            this.labelSkippedExisting.AutoSize = true;
            this.labelSkippedExisting.Location = new System.Drawing.Point(12, 241);
            this.labelSkippedExisting.Name = "labelSkippedExisting";
            this.labelSkippedExisting.Size = new System.Drawing.Size(194, 20);
            this.labelSkippedExisting.TabIndex = 15;
            this.labelSkippedExisting.Text = "Skipped Existing: 0 items.";
            //
            // labelReplacedExisting
            //
            this.labelReplacedExisting.AutoSize = true;
            this.labelReplacedExisting.Location = new System.Drawing.Point(278, 241);
            this.labelReplacedExisting.Name = "labelReplacedExisting";
            this.labelReplacedExisting.Size = new System.Drawing.Size(199, 20);
            this.labelReplacedExisting.TabIndex = 16;
            this.labelReplacedExisting.Text = "Replaced Existing: 0 items.";
            // 
            // ProgressForm
            // 
            this.ClientSize = new System.Drawing.Size(891, 802);
            this.Controls.Add(this.labelReplacedExisting);
            this.Controls.Add(this.labelSkippedExisting);
            this.Controls.Add(this.labelSkippedDate);
            this.Controls.Add(this.labelSkippedPermissions);
            this.Controls.Add(this.labelSkippedCast);
            this.Controls.Add(this.labelRetries);
            this.Controls.Add(this.labelSkippedItems);
            this.Controls.Add(this.labelDestination);
            this.Controls.Add(this.labelSource);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.labelLastError);
            this.Controls.Add(this.txtLastError);
            this.Controls.Add(this.labelErrorItems);
            this.Controls.Add(this.lblTimeRemaining);
            this.Controls.Add(this.lblTimeElapsed);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.progressBar);
            this.Name = "ProgressForm";
            this.Text = "Processing Items";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        private System.Windows.Forms.Label labelErrorItems;
        private System.Windows.Forms.Label labelLastError;
        private System.Windows.Forms.TextBox txtLastError;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Label labelSource;
        private System.Windows.Forms.Label labelDestination;
        private System.Windows.Forms.Label labelSkippedItems;
        private System.Windows.Forms.Label labelRetries;
        private System.Windows.Forms.Label labelSkippedCast;
        private System.Windows.Forms.Label labelSkippedPermissions;
        private System.Windows.Forms.Label labelSkippedDate;
        private System.Windows.Forms.Label labelSkippedExisting;
        private System.Windows.Forms.Label labelReplacedExisting;
    }
}
