using System;
using System.Windows.Forms;

namespace OutBack
{
    public partial class ProgressForm : Form
    {
        public ProgressForm()
        {
            InitializeComponent();
        }

        public void UpdateProgress(int processedItems, int totalItems, TimeSpan elapsedTime)
        {
            progressBar.Value = (int)((double)processedItems / totalItems * 100);
            lblStatus.Text = $"Processed {processedItems} of {totalItems} items.";
            lblTimeElapsed.Text = $"Time Elapsed: {elapsedTime.ToString(@"hh\:mm\:ss")}";
            TimeSpan estimatedRemaining = TimeSpan.FromTicks(elapsedTime.Ticks * (totalItems - processedItems) / processedItems);
            lblTimeRemaining.Text = $"Estimated Time Remaining: {estimatedRemaining.ToString(@"hh\:mm\:ss")}";
            Application.DoEvents();
        }
    }
}
