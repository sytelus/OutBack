namespace OutBack
{
    partial class MyRibbon : Microsoft.Office.Tools.Ribbon.OfficeRibbon
    {
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.groupOperations = this.Factory.CreateRibbonGroup();
            this.btnMoveCopy = this.Factory.CreateRibbonButton();
            //
            // tab1
            //
            this.tab1.Groups.Add(this.groupOperations);
            this.tab1.Label = "PST Mover";
            this.tab1.Name = "tab1";
            //
            // groupOperations
            //
            this.groupOperations.Items.Add(this.btnMoveCopy);
            this.groupOperations.Label = "Operations";
            this.groupOperations.Name = "groupOperations";
            //
            // btnMoveCopy
            //
            this.btnMoveCopy.Label = "Move/Copy to PST";
            this.btnMoveCopy.Name = "btnMoveCopy";
            this.btnMoveCopy.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMoveCopy_Click);
            //
            // MyRibbon
            //
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
        }

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupOperations;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMoveCopy;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibbon MyRibbon
        {
            get { return this.GetRibbon<MyRibbon>(); }
        }
    }
}
