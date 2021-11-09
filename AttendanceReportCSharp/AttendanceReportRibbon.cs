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

        int numDoorOpened = 0;
        int numRosterOpened = 0;
        ActionsPaneControl1 actionsPane1 = new ActionsPaneControl1();
        private void AttendanceReportRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            Globals.ThisWorkbook.ActionsPane.Controls.Add(actionsPane1);
            Globals.ThisWorkbook.Application.DisplayDocumentActionTaskPane = false;

            this.buttonOpenDoor.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
                this.buttonOpenDoor_Click);
            this.buttonOpenRoster.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
                this.buttonOpenRoster_Click);
            this.buttonOpenRoster.Enabled = false;

        }
        private void buttonOpenRoster_Click(object sender, RibbonControlEventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {

                Title = "Open Workbook",
                CheckFileExists = true,
                CheckPathExists = true,
                Filter = "Excel files (*.xls; *.xlsx)|*.xls;*.xlsx",
                RestoreDirectory = true,
                InitialDirectory = @"%USERPROFILE%\My Documents\Downloads",
                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Excel.Workbook rosterWB = Globals.ThisWorkbook.Application.Workbooks.Open(openFileDialog1.FileName, true, true);
                Excel.Worksheet rosterSheet1 = rosterWB.Worksheets[1];

                rosterSheet1.Copy(Globals.ThisWorkbook.Worksheets[1]);
                rosterWB.Close(false);

                Excel.Worksheet newRosterSheet = Globals.ThisWorkbook.Worksheets[1];
                newRosterSheet.Name = "Roster Report" + numRosterOpened.ToString();
                numRosterOpened++;

                Excel.Range lastRow = newRosterSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastUsedRow = lastRow.Row;

                Dictionary<String, int> dict = actionsPane1.numPerNameDict;
                for (int r = 1; r < lastUsedRow; r++)
                {
                    String name = newRosterSheet.Cells[r, 1].Value;
                    if (name != null)
                    {
                        name = name.ToLower();
                        name = name.Trim();
                        if (dict.ContainsKey(name))
                        {
                            newRosterSheet.Cells[r, 2] = dict[name];
                        } 
                        else
                        {
                            newRosterSheet.Cells[r, 2] = 0;
                        }
                    }
                }
                newRosterSheet.Range["A:B"].Sort(newRosterSheet.Columns[2]);

                int numDays = actionsPane1.Numdays;
                newRosterSheet.Range["E1"].Value2 = "Total Days";
                newRosterSheet.Range["E1"].Font.Bold = true;
                newRosterSheet.Range["E2"].Value2 = numDays.ToString();
                newRosterSheet.Range["E4"].Value2 = "Average Days in Studio";
                newRosterSheet.Range["E4"].Font.Bold = true;
                newRosterSheet.Range["E5"].Formula = "=AVERAGE(B1:B" + lastUsedRow + ")";
                newRosterSheet.Range["k1"].Value2 = "Less than 10% of time in-studio";
                newRosterSheet.Range["k1"].Font.Bold = true;
                newRosterSheet.Range["l1"].Formula = "=COUNTIF(B:B, \"<" + .1 * numDays + "\")";
                newRosterSheet.Range["k2"].Value2 = "Between 10% and 70% of time in-studio";
                newRosterSheet.Range["k2"].Font.Bold = true;
                newRosterSheet.Range["l2"].Formula = "=COUNTIFS(B:B, \">" + .1 * numDays + "\", B:B, \"<" + .69 * numDays + "\")";
                newRosterSheet.Range["k3"].Value2 = "Greater than 70% of time in-studio";
                newRosterSheet.Range["k3"].Font.Bold = true;
                newRosterSheet.Range["l3"].Formula = "=COUNTIF(B:B, \">=" + .7 * numDays + "\")";
                Excel.Range chartCells = newRosterSheet.Range["K1:K3", "L1:L3"];
                Excel.Range location = newRosterSheet.Range["E9"];
                location.Select();
                var chart = newRosterSheet.Shapes.AddChart2(-1, Microsoft.Office.Interop.Excel.XlChartType.xlPie, 300);
                chart.Chart.SetSourceData(chartCells);
                chart.Chart.ChartTitle.Text = "% Time In Studio";
                
                newRosterSheet.Range["A:M"].EntireColumn.AutoFit();




            }
        }
            private void buttonOpenDoor_Click(object sender, RibbonControlEventArgs e)
        {

            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {

                Title = "Open Workbook",
                CheckFileExists = true,
                CheckPathExists = true,
                Filter = "Excel files (*.xls; *.xlsx)|*.xls;*.xlsx",
                RestoreDirectory = true,
                InitialDirectory = @"%USERPROFILE%\My Documents\Downloads",
                ReadOnlyChecked = true,
                ShowReadOnly = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                CheckedListBox listbox = (CheckedListBox)actionsPane1.Controls["nameListAP"];
                listbox.Items.Clear();
                Excel.Workbook doorWB = Globals.ThisWorkbook.Application.Workbooks.Open(openFileDialog1.FileName, true, true);
                Excel.Worksheet doorSheet1 = doorWB.Worksheets[1];

                doorSheet1.Copy(Globals.ThisWorkbook.Worksheets[1]);
                doorWB.Close(false);

                Excel.Worksheet newDoorSheet = Globals.ThisWorkbook.Worksheets[1];
                newDoorSheet.Name = "Door Report" + numDoorOpened.ToString();
                numDoorOpened++;

                Excel.Range lastRow = newDoorSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int lastUsedRow = lastRow.Row;

                HashSet<String> nameHash = new HashSet<string>();
                for (int r = 7; r < lastUsedRow; r++)
                {
                    String name = newDoorSheet.Cells[r, 8].Value;
                    if (name != null)
                    {
                        name = name.ToLower();
                        name.Trim();
                        nameHash.Add(name);
                    }
                }
                List<String> nameList = new List<String>(nameHash);
                nameList.Sort();
                foreach (var item in nameList)
                {
                    listbox.Items.Add(item);
                }

                nameHash.Clear();

                Globals.ThisWorkbook.Application.DisplayDocumentActionTaskPane = true;
            }
        }
    }
}
