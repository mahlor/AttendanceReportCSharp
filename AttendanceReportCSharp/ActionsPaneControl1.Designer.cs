namespace AttendanceReportCSharp
{
    [System.ComponentModel.ToolboxItemAttribute(false)]
    partial class ActionsPaneControl1
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
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
            this.label1 = new System.Windows.Forms.Label();
            this.nameListAP = new System.Windows.Forms.CheckedListBox();
            this.removeDupsAPButton = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.checkboxSelectAll = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 25);
            this.label1.TabIndex = 0;
            // 
            // nameListAP
            // 
            this.nameListAP.BackColor = System.Drawing.Color.White;
            this.nameListAP.FormattingEnabled = true;
            this.nameListAP.Location = new System.Drawing.Point(6, 213);
            this.nameListAP.Name = "nameListAP";
            this.nameListAP.Size = new System.Drawing.Size(724, 788);
            this.nameListAP.TabIndex = 1;
            // 
            // removeDupsAPButton
            // 
            this.removeDupsAPButton.AutoSize = true;
            this.removeDupsAPButton.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.removeDupsAPButton.Location = new System.Drawing.Point(9, 10);
            this.removeDupsAPButton.Name = "removeDupsAPButton";
            this.removeDupsAPButton.Size = new System.Drawing.Size(208, 35);
            this.removeDupsAPButton.TabIndex = 2;
            this.removeDupsAPButton.Text = "Remove Duplicates";
            this.removeDupsAPButton.UseVisualStyleBackColor = true;
            this.removeDupsAPButton.Click += new System.EventHandler(this.removeDupsAPButton_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 58);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(274, 50);
            this.label2.TabIndex = 3;
            this.label2.Text = "Check any names you wish\r\nto remove from results";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // checkboxSelectAll
            // 
            this.checkboxSelectAll.AutoSize = true;
            this.checkboxSelectAll.Location = new System.Drawing.Point(9, 178);
            this.checkboxSelectAll.Name = "checkboxSelectAll";
            this.checkboxSelectAll.Size = new System.Drawing.Size(191, 29);
            this.checkboxSelectAll.TabIndex = 4;
            this.checkboxSelectAll.Text = "Select All/None";
            this.checkboxSelectAll.UseVisualStyleBackColor = true;
            this.checkboxSelectAll.CheckedChanged += new System.EventHandler(this.selectall_CheckedChanged);
            // 
            // ActionsPaneControl1
            // 
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Controls.Add(this.checkboxSelectAll);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.removeDupsAPButton);
            this.Controls.Add(this.nameListAP);
            this.Controls.Add(this.label1);
            this.Name = "ActionsPaneControl1";
            this.Size = new System.Drawing.Size(733, 1004);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckedListBox nameListAP;
        private System.Windows.Forms.Button removeDupsAPButton;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox checkboxSelectAll;
    }
}
