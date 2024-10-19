using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutBack
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnMoveCopy_Click(object sender, RibbonControlEventArgs e)
        {
            PSTSelectionForm pstForm = new PSTSelectionForm();
            if (pstForm.ShowDialog() == DialogResult.OK)
            {
                var pstFilePath = pstForm.PstFilePath;
                var isMoveOperation = pstForm.IsMoveOperation;

                ItemMover mover = new ItemMover();
                mover.Start(Globals.ThisAddIn.Application.ActiveExplorer().CurrentFolder, pstFilePath, isMoveOperation);
            }
        }
    }
}
