using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace AttendanceReportCSharp
{
    partial class ActionsPaneControl1 : UserControl
    {
        int numDays = 0;
        int numPerDay = 0;
        int numOpened = 0;
        Dictionary<DateTime, int> numPerDayDict = new Dictionary<DateTime, int> { };
        Dictionary<String, int> numPerNameDict = new Dictionary<String, int> { };

        public ActionsPaneControl1()
        {
            InitializeComponent();
        }

        private void removeDupsAPButton_Click(object sender, EventArgs e)
        {
            Excel.Worksheet activesheet = Globals.ThisWorkbook.Application.ActiveSheet;
            if (activesheet.Name.StartsWith("Remove")) return;

            HashSet<String> nameHash = new HashSet<string>();
            List<String> exceptions = new List<String>();
            foreach (String item in nameListAP.CheckedItems)
            {
                exceptions.Add(item);
            }

            Excel.Worksheet newDoorSheet = Globals.ThisWorkbook.Worksheets[1];
            newDoorSheet.Copy(Globals.ThisWorkbook.Worksheets[1]);

            Excel.Worksheet removeDupsSheet = Globals.ThisWorkbook.Worksheets[1];
            removeDupsSheet.Name = "Remove Dups" + numOpened.ToString();
            numOpened++;
            removeDupsSheet.Range["A1:A6"].EntireRow.Delete();
            removeDupsSheet.Range["A1"].EntireColumn.Delete();
            removeDupsSheet.Range["B1:F1"].EntireColumn.Delete();
            removeDupsSheet.Range["C1:G1"].EntireColumn.Delete();
            Excel.Range lastRow = removeDupsSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastUsedRow = lastRow.Row;

            for (int r = 1; r < lastUsedRow - 1; r++)
            {
                String name = removeDupsSheet.Cells[r, 2].Value;
                String dateStr = removeDupsSheet.Cells[r, 1].Value;
                if (name != null)
                {
                    name = name.ToLower();
                }
                if (dateStr != null)
                {
                    dateStr = dateStr.Split(' ')[0];
                }
                removeDupsSheet.Range["C" + r].Value2 = dateStr;
                removeDupsSheet.Range["D" + r].Value2 = name;
            }
            removeDupsSheet.Range["A1:B1"].EntireColumn.Delete();

            DateTime dateMatch = removeDupsSheet.Cells[1, 1].Value;
            numDays++;
            DateTime dateOrg;
            var names = new List<(DateTime dateList, string nameList)> { };
            for (int r = 1; r < lastUsedRow; r++)
            {
                if (removeDupsSheet.Cells[r, 1].Value != null && removeDupsSheet.Cells[r, 2].Value != null)
                {
                    dateOrg = removeDupsSheet.Cells[r, 1].Value;
                    if (dateOrg.DayOfWeek != DayOfWeek.Saturday && dateOrg.DayOfWeek != DayOfWeek.Sunday)
                    {
                        if (dateOrg != dateMatch)
                        {

                            foreach (String name in nameHash)
                            {
                                names.Add((dateMatch, name));
                                if (numPerNameDict.ContainsKey(name))
                                {
                                    numPerNameDict[name] = numPerNameDict[name] + 1;
                                } 
                                else
                                {
                                    numPerNameDict.Add(name, 1);
                                }

                            }
                            numPerDay = nameHash.Count;
                            numPerDayDict.Add(dateMatch, numPerDay);
                            dateMatch = dateOrg;
                            numDays++;
                            nameHash.Clear();
                            String item = removeDupsSheet.Cells[r, 2].value;
                            item.Trim();
                            if (!exceptions.Contains(item))
                            {
                                nameHash.Add(item);
                            }

                        }
                        else
                        {
                            String item = removeDupsSheet.Cells[r, 2].value;
                            item.Trim();
                            if (!exceptions.Contains(item))
                            {
                                nameHash.Add(item);
                            }

                        }
                    }

                }
            }
            for (int s = 1, t = 0; t < names.Count; s++, t++)
            {
                removeDupsSheet.Range["C" + s].Value = names[t].dateList;
                removeDupsSheet.Range["D" + s].Value = names[t].nameList;

            }
            removeDupsSheet.Range["A1:B1"].EntireColumn.Delete();

            int cell = 9;
            foreach (var day in numPerDayDict)
            {
                removeDupsSheet.Range["E" + cell].Value2 = day.Key.ToShortDateString();
                removeDupsSheet.Range["F" + cell].Value2 = day.Value.ToString();
                cell++;
            }
            removeDupsSheet.Range["E1"].Value2 = "Total Days";
            removeDupsSheet.Range["E2"].Value2 = numDays.ToString();
            removeDupsSheet.Range["E4"].Value2 = "Average Per Day";
            removeDupsSheet.Range["E5"].Formula = "=AVERAGE(F9:F" + (cell - 1) + ")";
            removeDupsSheet.Range["E7"].Value2 = "Total in Each Day";

            cell = 5;

            foreach (var name in numPerNameDict)
            {
                removeDupsSheet.Range["H" + cell].Value2 = name.Key;
                removeDupsSheet.Range["I" + cell].Value2 = name.Value.ToString();
                cell++;
            }

            numPerDayDict.Clear();
            numPerNameDict.Clear();
            removeDupsSheet.Range["H:I"].Sort(removeDupsSheet.Columns[9]);
            removeDupsSheet.Range["A1:M1"].EntireColumn.AutoFit();

            Globals.ThisWorkbook.Application.DisplayDocumentActionTaskPane = false;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void nameListAP_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
