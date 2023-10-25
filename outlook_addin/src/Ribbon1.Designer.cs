namespace Outlook2Aula
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.tabO2A = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnRunO2A = this.Factory.CreateRibbonButton();
            this.btnForceUpdate = this.Factory.CreateRibbonButton();
            this.grpSettings = this.Factory.CreateRibbonGroup();
            this.btnSelectO2AFolder = this.Factory.CreateRibbonButton();
            this.btnAllSettings = this.Factory.CreateRibbonButton();
            this.btnOpenIgnoreFile = this.Factory.CreateRibbonButton();
            this.btnOpenPeopleWorkbook = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.box2 = this.Factory.CreateRibbonBox();
            this.label2 = this.Factory.CreateRibbonLabel();
            this.lblO2APath = this.Factory.CreateRibbonLabel();
            this.box1 = this.Factory.CreateRibbonBox();
            this.label1 = this.Factory.CreateRibbonLabel();
            this.lblBuildVersion = this.Factory.CreateRibbonLabel();
            this.tabO2A.SuspendLayout();
            this.group1.SuspendLayout();
            this.grpSettings.SuspendLayout();
            this.group3.SuspendLayout();
            this.box2.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabO2A
            // 
            this.tabO2A.Groups.Add(this.group1);
            this.tabO2A.Groups.Add(this.grpSettings);
            this.tabO2A.Groups.Add(this.group3);
            this.tabO2A.Label = "Outlook2Aula";
            this.tabO2A.Name = "tabO2A";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnRunO2A);
            this.group1.Items.Add(this.btnForceUpdate);
            this.group1.Label = "Afvikling";
            this.group1.Name = "group1";
            // 
            // btnRunO2A
            // 
            this.btnRunO2A.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRunO2A.Image = global::Outlook2Aula.Properties.Resources.exchange;
            this.btnRunO2A.Label = "Kør";
            this.btnRunO2A.Name = "btnRunO2A";
            this.btnRunO2A.ShowImage = true;
            this.btnRunO2A.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // btnForceUpdate
            // 
            this.btnForceUpdate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnForceUpdate.Image = global::Outlook2Aula.Properties.Resources.exchange;
            this.btnForceUpdate.Label = "Kør og opdater alt (Force update)";
            this.btnForceUpdate.Name = "btnForceUpdate";
            this.btnForceUpdate.ShowImage = true;
            this.btnForceUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnForceUpdate_Click);
            // 
            // grpSettings
            // 
            this.grpSettings.Items.Add(this.btnSelectO2AFolder);
            this.grpSettings.Items.Add(this.btnAllSettings);
            this.grpSettings.Items.Add(this.btnOpenIgnoreFile);
            this.grpSettings.Items.Add(this.btnOpenPeopleWorkbook);
            this.grpSettings.Label = "O2A Indstillinger";
            this.grpSettings.Name = "grpSettings";
            // 
            // btnSelectO2AFolder
            // 
            this.btnSelectO2AFolder.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSelectO2AFolder.Image = global::Outlook2Aula.Properties.Resources.settings;
            this.btnSelectO2AFolder.Label = "Vælg O2A mappe";
            this.btnSelectO2AFolder.Name = "btnSelectO2AFolder";
            this.btnSelectO2AFolder.ScreenTip = "Vælg hvor O2A mappen befinder sig. Du skal vælge mappen hvori f.eks. \"src\" mappen" +
    ", \"personer.csv\" og \"personer_ignorer.csv\" filerne findes";
            this.btnSelectO2AFolder.ShowImage = true;
            this.btnSelectO2AFolder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectO2AFolder_Click);
            // 
            // btnAllSettings
            // 
            this.btnAllSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAllSettings.Image = global::Outlook2Aula.Properties.Resources.settings;
            this.btnAllSettings.Label = "Åben O2A indstillinger";
            this.btnAllSettings.Name = "btnAllSettings";
            this.btnAllSettings.ShowImage = true;
            this.btnAllSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAllSettings_Click);
            // 
            // btnOpenIgnoreFile
            // 
            this.btnOpenIgnoreFile.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnOpenIgnoreFile.Image = global::Outlook2Aula.Properties.Resources.ignore;
            this.btnOpenIgnoreFile.Label = "Administer ignorede personer (Regneark)";
            this.btnOpenIgnoreFile.Name = "btnOpenIgnoreFile";
            this.btnOpenIgnoreFile.ShowImage = true;
            this.btnOpenIgnoreFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpenIgnoreFile_Click);
            // 
            // btnOpenPeopleWorkbook
            // 
            this.btnOpenPeopleWorkbook.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnOpenPeopleWorkbook.Image = global::Outlook2Aula.Properties.Resources._switch;
            this.btnOpenPeopleWorkbook.Label = "Administer personer (Regneark)";
            this.btnOpenPeopleWorkbook.Name = "btnOpenPeopleWorkbook";
            this.btnOpenPeopleWorkbook.ShowImage = true;
            this.btnOpenPeopleWorkbook.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOpenPeopleWorkbook_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.box2);
            this.group3.Items.Add(this.box1);
            this.group3.Label = "AddIn oplysninger";
            this.group3.Name = "group3";
            // 
            // box2
            // 
            this.box2.Items.Add(this.label2);
            this.box2.Items.Add(this.lblO2APath);
            this.box2.Name = "box2";
            // 
            // label2
            // 
            this.label2.Label = "O2A sti:";
            this.label2.Name = "label2";
            // 
            // lblO2APath
            // 
            this.lblO2APath.Label = "UKENDT STI";
            this.lblO2APath.Name = "lblO2APath";
            // 
            // box1
            // 
            this.box1.Items.Add(this.label1);
            this.box1.Items.Add(this.lblBuildVersion);
            this.box1.Name = "box1";
            // 
            // label1
            // 
            this.label1.Label = "Version:";
            this.label1.Name = "label1";
            // 
            // lblBuildVersion
            // 
            this.lblBuildVersion.Label = "Version";
            this.lblBuildVersion.Name = "lblBuildVersion";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tabO2A);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tabO2A.ResumeLayout(false);
            this.tabO2A.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.grpSettings.ResumeLayout(false);
            this.grpSettings.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabO2A;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAllSettings;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRunO2A;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnForceUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectO2AFolder;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblO2APath;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpenIgnoreFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOpenPeopleWorkbook;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblBuildVersion;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel label1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
