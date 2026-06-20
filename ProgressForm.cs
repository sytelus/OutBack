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

        public void SetCurrentFolders(string sourceFolder, string destinationFolder)
        {
            labelSource.Text = $"Source: {sourceFolder}";
            labelDestination.Text = $"Destination: {destinationFolder}";
            Application.DoEvents();
        }

        public bool UpdateProgress(int processedItems, int totalItems, int errorItems, int skippedItems,
            int skipForCast, int skipForPermission, int skipForDate, int skipForExisting, int replacedExisting,
            string lastError, TimeSpan elapsedTime, int retryCount)
        {
            int progressValue = totalItems <= 0 ? 100 : Math.Min((int)((double)processedItems / totalItems * 100), 100);
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
            labelSkippedExisting.Text = $"Skipped Existing: {skipForExisting} items.";
            labelReplacedExisting.Text = $"Replaced Existing: {replacedExisting} items.";
            labelRetries.Text = $"Retries: {retryCount}";
            if (!string.IsNullOrEmpty(lastError))
            {
                txtLastError.Text = lastError;
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
