using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutBack
{
    public partial class MyRibbon : OfficeRibbon
    {
        public MyRibbon()
        {
            InitializeComponent();
        }

        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            // Initialization code if needed
        }

        private void btnMoveCopy_Click(object sender, RibbonControlEventArgs e)
        {
            PSTSelectionForm pstForm = new PSTSelectionForm();
            if (pstForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var pstFilePath = pstForm.PstFilePath;
                var isMoveOperation = pstForm.IsMoveOperation;

                ItemMover mover = new ItemMover();
                mover.Start(Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder, pstFilePath, isMoveOperation);
            }
        }
    }
}
