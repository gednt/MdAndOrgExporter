namespace MdAndOrgExporter
{
    partial class Export : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Export()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnExportToMd = this.Factory.CreateRibbonButton();
            this.ExportToOrg = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.txtTags = this.Factory.CreateRibbonEditBox();
            this.chkSeparateTags = this.Factory.CreateRibbonCheckBox();
            this.chkAllowImages = this.Factory.CreateRibbonCheckBox();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnExportToMd);
            this.group1.Items.Add(this.ExportToOrg);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.txtTags);
            this.group1.Items.Add(this.chkSeparateTags);
            this.group1.Items.Add(this.chkAllowImages);
            this.group1.Label = "Utilities";
            this.group1.Name = "group1";
            // 
            // btnExportToMd
            // 
            this.btnExportToMd.Label = "Export to MD";
            this.btnExportToMd.Name = "btnExportToMd";
            this.btnExportToMd.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportToMd_Click);
            // 
            // ExportToOrg
            // 
            this.ExportToOrg.Label = "Export to Org";
            this.ExportToOrg.Name = "ExportToOrg";
            this.ExportToOrg.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportToOrg_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // txtTags
            // 
            this.txtTags.Label = "Tags";
            this.txtTags.Name = "txtTags";
            this.txtTags.Text = null;
            // 
            // chkSeparateTags
            // 
            this.chkSeparateTags.Label = "Separate Tags";
            this.chkSeparateTags.Name = "chkSeparateTags";
            // 
            // chkAllowImages
            // 
            this.chkAllowImages.Checked = true;
            this.chkAllowImages.Label = "Allow Images";
            this.chkAllowImages.Name = "chkAllowImages";
            // 
            // Export
            // 
            this.Name = "Export";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Export_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ExportToOrg;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox txtTags;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportToMd;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkSeparateTags;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkAllowImages;
    }

    partial class ThisRibbonCollection
    {
        internal Export Export
        {
            get { return this.GetRibbon<Export>(); }
        }
    }
}
