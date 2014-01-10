namespace Simple_Signature
{
    partial class RibbonMail : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonMail()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur de composants

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonMail));
            this.group1 = this.Factory.CreateRibbonGroup();
            this.SignatureGallery = this.Factory.CreateRibbonGallery();
            tab1 = this.Factory.CreateRibbonTab();
            tab1.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            tab1.Groups.Add(this.group1);
            tab1.Label = "Simple Signature";
            tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.SignatureGallery);
            this.group1.Name = "group1";
            // 
            // SignatureGallery
            // 
            this.SignatureGallery.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SignatureGallery.Image = ((System.Drawing.Image)(resources.GetObject("SignatureGallery.Image")));
            this.SignatureGallery.Label = "Signatures";
            this.SignatureGallery.Name = "SignatureGallery";
            this.SignatureGallery.ShowImage = true;
            // 
            // RibbonMail
            // 
            this.Name = "RibbonMail";
            this.RibbonType = "Microsoft.Outlook.Mail.Compose, Microsoft.Outlook.Response.Compose";
            this.Tabs.Add(tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonMail_Load);
            tab1.ResumeLayout(false);
            tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery SignatureGallery;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonMail RibbonMail
        {
            get { return this.GetRibbon<RibbonMail>(); }
        }
    }
}
