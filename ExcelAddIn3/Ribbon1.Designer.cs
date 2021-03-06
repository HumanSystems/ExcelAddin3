﻿namespace ExcelAddIn2
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
            this.tabSA = this.Factory.CreateRibbonTab();
            this.grpLotMgt = this.Factory.CreateRibbonGroup();
            this.btnLoadCatMast = this.Factory.CreateRibbonButton();
            this.btnVerify = this.Factory.CreateRibbonButton();
            this.btnSelectCategory = this.Factory.CreateRibbonButton();
            this.tabSA.SuspendLayout();
            this.grpLotMgt.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabSA
            // 
            this.tabSA.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tabSA.Groups.Add(this.grpLotMgt);
            this.tabSA.Label = "AMS Functions";
            this.tabSA.Name = "tabSA";
            // 
            // grpLotMgt
            // 
            this.grpLotMgt.Items.Add(this.btnLoadCatMast);
            this.grpLotMgt.Items.Add(this.btnVerify);
            this.grpLotMgt.Items.Add(this.btnSelectCategory);
            this.grpLotMgt.Label = "Lot Mgt";
            this.grpLotMgt.Name = "grpLotMgt";
            // 
            // btnLoadCatMast
            // 
            this.btnLoadCatMast.Label = "1) (Re)Create AMS Spreadsheet from CM Spreadsheet";
            this.btnLoadCatMast.Name = "btnLoadCatMast";
            this.btnLoadCatMast.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnLoadCatMast_Click);
            // 
            // btnVerify
            // 
            this.btnVerify.Label = "2) (ReFuck)Verify Data";
            this.btnVerify.Name = "btnVerify";
            this.btnVerify.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnVerify_Click_1);
            // 
            // btnSelectCategory
            // 
            this.btnSelectCategory.Label = "3) Select/Populate category for rows in filter";
            this.btnSelectCategory.Name = "btnSelectCategory";
            //this.btnSelectCategory.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelectCategory_Click_1);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tabSA);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tabSA.ResumeLayout(false);
            this.tabSA.PerformLayout();
            this.grpLotMgt.ResumeLayout(false);
            this.grpLotMgt.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabSA;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpLotMgt;
        //internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadAMS;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnLoadCatMast;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnVerify;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSelectCategory;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}