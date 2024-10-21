//using Microsoft.Office.Tools.Ribbon;

namespace OutBack
{
    public partial class ThisAddIn
    {
        //private MyRibbon ribbon;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Initialization code if needed
        }

        //protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        //{
        //    ribbon = new MyRibbon();
        //    return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new IRibbonExtension[] { ribbon });
        //}

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Cleanup code if needed
        }

        // No need to override CreateRibbonExtensibilityObject when using the designer

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
