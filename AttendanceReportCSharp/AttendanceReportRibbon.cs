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
         //   this.buttonRemoveDups.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
           //     this.buttonRemoveDups_Click);
            this.buttonOpenRoster.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
                this.buttonOpenRoster_Click);
            this.buttonOpenRoster.Enabled = false;
            //            this.buttonCalcDays.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(
            //              this.buttonCalcDays_Click);

        }
        /*        private void buttonCalcDays_Click(object sender, RibbonControlEventArgs e)
                {
                    //  Globals.ThisWorkbook.Application.DisplayDocumentActionTaskPane = true;
                    //actionsPane1.Show();
                }
        */

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
                //Excel.Worksheet currentWS = (Excel.Worksheet)Globals.ThisWorkbook.Application.ActiveWorkbook.ActiveSheet;
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
        private void buttonRemoveDups_Click(object sender, RibbonControlEventArgs e)
        {
     //       Globals.ThisWorkbook.Application.DisplayDocumentActionTaskPane = true;
       //     actionsPane1.Show();
        }
        
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void buttonOpenDoor_Click_1(object sender, RibbonControlEventArgs e)
        {

        }
    }
}
