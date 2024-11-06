using System;
using System.Windows.Forms;

namespace OutBack
{
    public partial class ProgressForm : Form
    {
        bool isCancelled = false;

        public ProgressForm(string sourceFolder, string destinationFolder)
        {
            InitializeComponent();
            labelSource.Text = $"Source: {sourceFolder}";
            labelDestination.Text = $"Destination: {destinationFolder}";
            labelRetries.Text = "Retries: 0";
        }

        public bool UpdateProgress(int processedItems, int totalItems, int errorItems, int skippedItems,
            int skipForCast, int skipForPermission, int skipForDate,
            string lastError, TimeSpan elapsedTime, int retryCount)
        {
            int progressValue = Math.Min((int)((double)processedItems / totalItems * 100), 100);
            progressBar.Value = progressValue;

            lblStatus.Text = $"Processed {processedItems} of {totalItems} items.";
            lblTimeElapsed.Text = $"Time Elapsed: {elapsedTime.ToString(@"hh\:mm\:ss")}";

            TimeSpan estimatedRemaining = TimeSpan.Zero;
            if (processedItems > 0)
            {
                estimatedRemaining = TimeSpan.FromTicks(elapsedTime.Ticks * (totalItems - processedItems) / processedItems);
            }
            lblTimeRemaining.Text = $"Estimated Time Remaining: {estimatedRemaining.ToString(@"hh\:mm\:ss")}";

            labelErrorItems.Text = $"Errors: {errorItems} items.";
            labelSkippedItems.Text = $"Skipped Total: {skippedItems} items.";
            labelSkippedCast.Text = $"Skipped Cast: {skipForCast} items.";
            labelSkippedPermissions.Text = $"Skipped Perms: {skipForPermission} items.";
            labelSkippedDate.Text = $"Skipped Date: {skipForDate} items.";
            labelRetries.Text = $"Retries: {retryCount}";
            if (!string.IsNullOrEmpty(lastError))
            {
                labelLastError.Text = $"Last Error: '{lastError}'";
            }
            Application.DoEvents();
            return isCancelled;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            isCancelled = true;
        }
    }
}
