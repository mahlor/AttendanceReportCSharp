using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace AttendanceReportCSharp
{
    public partial class AttendanceReportRibbon
    {
        ActionsPaneControl1 actionsPane1 = new ActionsPaneControl1();
        private void AttendanceReportRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.ThisWorkbook.ActionsPane.Controls.Add(actionsPane1);
            actionsPane1.Hide();
            Globals.ThisWorkbook.Application.DisplayDocumentActionTaskPane = false;

            this.buttonOpenDoor.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
                this.buttonOpenDoor_Click);
            this.buttonRemoveDups.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
                this.buttonRemoveDups_Click);
            this.buttonOpenRoster.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
                this.buttonOpenRoster_Click);
            this.buttonCalcDays.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
                this.buttonCalcDays_Click);

        }
        private void buttonCalcDays_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Application.DisplayDocumentActionTaskPane = true;
            actionsPane1.Show();

        }

        private void buttonRemoveDups_Click(object sender, RibbonControlEventArgs e)
        {

        }
        private void buttonOpenDoor_Click(object sender, RibbonControlEventArgs e)
        {
            BrowseButton_Click(sender, e);

        }
        private void buttonOpenRoster_Click(object sender, RibbonControlEventArgs e)
        {
            BrowseButton_Click(sender, e);

        }
        private void BrowseButton_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {

                Title = "Open Workbook",
                CheckFileExists = true,
                CheckPathExists = true,
                Filter = "Excel files (*.xls; *.xlsx)|*.xls;*.xlsx",
                RestoreDirectory = true,

                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Excel.Worksheet currentWS = (Excel.Worksheet)Globals.ThisWorkbook.Application.ActiveWorkbook.ActiveSheet;

                var writeRange = currentWS.Range["A2"];
                writeRange.Value = openFileDialog1.FileName;





































            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
