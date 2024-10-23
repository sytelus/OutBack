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
            this.buttonCancel = new System.Windows.Forms.Button();
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
            this.labelErrorItems.Location = new System.Drawing.Point(12, 147);
            this.labelErrorItems.Name = "labelErrorItems";
            this.labelErrorItems.Size = new System.Drawing.Size(115, 20);
            this.labelErrorItems.TabIndex = 4;
            this.labelErrorItems.Text = "Errors: 0 items.";
            // 
            // labelLastError
            // 
            this.labelLastError.AutoSize = true;
            this.labelLastError.Location = new System.Drawing.Point(12, 185);
            this.labelLastError.Name = "labelLastError";
            this.labelLastError.Size = new System.Drawing.Size(82, 40);
            this.labelLastError.TabIndex = 5;
            this.labelLastError.Text = "Error\r\nMessages\r\n";
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
            // ProgressForm
            // 
            this.ClientSize = new System.Drawing.Size(740, 542);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.labelLastError);
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
        private System.Windows.Forms.Button buttonCancel;
    }
}
