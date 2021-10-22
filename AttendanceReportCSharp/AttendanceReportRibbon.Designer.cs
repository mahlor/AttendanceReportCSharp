namespace AttendanceReportCSharp
{
    partial class AttendanceReportRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AttendanceReportRibbon()
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
            this.buttonRemoveDups = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.buttonOpenDoor = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.buttonOpenRoster = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.buttonCalcDays = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Label = "Attendance Report";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.buttonRemoveDups);
            this.group1.Name = "group1";
            // 
            // buttonRemoveDups
            // 
            this.buttonRemoveDups.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonRemoveDups.Label = "Remove Duplicates";
            this.buttonRemoveDups.Name = "buttonRemoveDups";
            this.buttonRemoveDups.OfficeImageId = "ReplaceDialog";
            this.buttonRemoveDups.ShowImage = true;
            // 
            // group2
            // 
            this.group2.Items.Add(this.buttonOpenDoor);
            this.group2.Name = "group2";
            // 
            // buttonOpenDoor
            // 
            this.buttonOpenDoor.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonOpenDoor.Label = "Open Door Report ";
            this.buttonOpenDoor.Name = "buttonOpenDoor";
            this.buttonOpenDoor.OfficeImageId = "FileOpen";
            this.buttonOpenDoor.ShowImage = true;
            // 
            // group3
            // 
            this.group3.Items.Add(this.buttonOpenRoster);
            this.group3.Name = "group3";
            // 
            // buttonOpenRoster
            // 
            this.buttonOpenRoster.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonOpenRoster.Label = "Open Studio Roster";
            this.buttonOpenRoster.Name = "buttonOpenRoster";
            this.buttonOpenRoster.OfficeImageId = "FileOpen";
            this.buttonOpenRoster.ShowImage = true;
            this.buttonOpenRoster.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.buttonCalcDays);
            this.group4.Name = "group4";
            // 
            // buttonCalcDays
            // 
            this.buttonCalcDays.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonCalcDays.Label = "Calculate Days in Studio";
            this.buttonCalcDays.Name = "buttonCalcDays";
            this.buttonCalcDays.OfficeImageId = "AutoSum";
            this.buttonCalcDays.ShowImage = true;
            // 
            // AttendanceReportRibbon
            // 
            this.Name = "AttendanceReportRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AttendanceReportRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRemoveDups;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonOpenDoor;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonOpenRoster;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCalcDays;
    }

    partial class ThisRibbonCollection
    {
        internal AttendanceReportRibbon AttendanceReportRibbon
        {
            get { return this.GetRibbon<AttendanceReportRibbon>(); }
        }
    }
}
