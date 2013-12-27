﻿namespace Simple_Signature
{
    partial class RibbonExplorer : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonExplorer()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonExplorer));
            this.SimpleSignTab = this.Factory.CreateRibbonTab();
            this.OptionGroup = this.Factory.CreateRibbonGroup();
            this.showOptions = this.Factory.CreateRibbonButton();
            this.showAdvancedOptions = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.SignatureGallery = this.Factory.CreateRibbonGallery();
            this.SimpleSignTab.SuspendLayout();
            this.OptionGroup.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // SimpleSignTab
            // 
            this.SimpleSignTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.SimpleSignTab.Groups.Add(this.OptionGroup);
            this.SimpleSignTab.Groups.Add(this.group1);
            this.SimpleSignTab.Label = "Simple Sign";
            this.SimpleSignTab.Name = "SimpleSignTab";
            // 
            // OptionGroup
            // 
            this.OptionGroup.Items.Add(this.showOptions);
            this.OptionGroup.Items.Add(this.showAdvancedOptions);
            this.OptionGroup.Name = "OptionGroup";
            // 
            // showOptions
            // 
            this.showOptions.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.showOptions.Image = ((System.Drawing.Image)(resources.GetObject("showOptions.Image")));
            this.showOptions.Label = "Options";
            this.showOptions.Name = "showOptions";
            this.showOptions.ShowImage = true;
            // 
            // showAdvancedOptions
            // 
            this.showAdvancedOptions.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.showAdvancedOptions.Image = ((System.Drawing.Image)(resources.GetObject("showAdvancedOptions.Image")));
            this.showAdvancedOptions.Label = "Options Avancées";
            this.showAdvancedOptions.Name = "showAdvancedOptions";
            this.showAdvancedOptions.ShowImage = true;
            // 
            // group1
            // 
            this.group1.Items.Add(this.SignatureGallery);
            this.group1.Name = "group1";
            // 
            // SignatureGallery
            // 
            this.SignatureGallery.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.SignatureGallery.Label = "Signatures";
            this.SignatureGallery.Name = "SignatureGallery";
            this.SignatureGallery.ShowImage = true;
            // 
            // RibbonExplorer
            // 
            this.Name = "RibbonExplorer";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.SimpleSignTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonExplorer_Load);
            this.SimpleSignTab.ResumeLayout(false);
            this.SimpleSignTab.PerformLayout();
            this.OptionGroup.ResumeLayout(false);
            this.OptionGroup.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab SimpleSignTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup OptionGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton showOptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton showAdvancedOptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery SignatureGallery;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonExplorer RibbonExplorer
        {
            get { return this.GetRibbon<RibbonExplorer>(); }
        }
    }
}
