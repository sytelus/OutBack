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
        }

        public bool UpdateProgress(int processedItems, int totalItems, int errorItems, string lastError, TimeSpan elapsedTime)
        {
            progressBar.Value = (int)((double)processedItems / totalItems * 100);
            lblStatus.Text = $"Processed {processedItems} of {totalItems} items.";
            lblTimeElapsed.Text = $"Time Elapsed: {elapsedTime.ToString(@"hh\:mm\:ss")}";
            TimeSpan estimatedRemaining = TimeSpan.FromTicks(elapsedTime.Ticks * (totalItems - processedItems) / processedItems);
            lblTimeRemaining.Text = $"Estimated Time Remaining: {estimatedRemaining.ToString(@"hh\:mm\:ss")}";
            labelErrorItems.Text = $"Error Items: {errorItems}";
            if (!string.IsNullOrEmpty(lastError))
            {
                labelErrorItems.Text = $"\nLast Error: '{lastError}'";
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
