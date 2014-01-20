using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Diagnostics;

namespace JDTReport
{
    public partial class Sheet2
    {
        private void Sheet2_Startup(object sender, System.EventArgs e)
        {
        }

        private void Sheet2_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.button1.Click += new System.EventHandler(this.button1_Click);
            this.Startup += new System.EventHandler(this.Sheet2_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet2_Shutdown);

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            // Selects the data sheet that contains the RAW data from JDT          
            Excel.Worksheet RawData_Sheet = (Excel.Worksheet)this.Application.Worksheets["Raw Data"];

            RawData_Sheet.Select(true);

            // Grabs range for the test end dates for further use
            Excel.Range TestEnd_Range = RawData_Sheet.get_Range("O1", "O" + RawData_Sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row);

            // Grabs range of report write dates for further use
            Excel.Range ReportWrite_Range = RawData_Sheet.get_Range("Q1", "Q" + RawData_Sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row);

            // Grabs range of JobNumbers for further use -- Unused currently, but good to have I suppose.
            Excel.Range JobNumber_Range = RawData_Sheet.get_Range("A1", "A" + RawData_Sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row);

            // Method do delete all columns that do are NOT used in the following analysis (caliper information, etc.)
            Delete_Unused_Columns(RawData_Sheet);
            
            // Deletes all rows which do NOT have a test end date (tests still in progress).
            Delete_Empty_Rows(RawData_Sheet, TestEnd_Range);

            // Manually adds calculations for both the test write and check times
            Add_Metric_Calculations(RawData_Sheet);

            Excel.Range TimetoWrite_Range = RawData_Sheet.get_Range("K1", "K" + RawData_Sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row);

            Delete_Invalid_Rows(RawData_Sheet, TimetoWrite_Range);

            // Creates Excel.Range for entirety of filtered results
            Excel.Range Filtered_Data_Range = Filter_By_Date(RawData_Sheet, TestEnd_Range);

            Paste_Filtered_Data(Filtered_Data_Range);

            Create_Pivot_Table(this.Application.Worksheets["Pivot Table"]);
        }

        public void Create_Pivot_Table(Excel.Worksheet Pivot_Table_WorkSheet)
        {
            Excel.Range Pivot_Table_Range = Pivot_Table_WorkSheet.UsedRange;

            Excel.PivotCache pivot_Cache = this.Application.ActiveWorkbook.PivotCaches().Add(Excel.XlPivotTableSourceType.xlDatabase, Pivot_Table_Range);

            Excel.PivotTable pivotTable = pivot_Cache.CreatePivotTable(Pivot_Table_WorkSheet.Range["P2"], "Table Information", Pivot_Table_Range);

            
        }

        public void Paste_Filtered_Data(Excel.Range Filtered_Data_Range)
        {
            Filtered_Data_Range.Copy();

            Excel.Worksheet PivotTable_WorkSheet = this.Application.ActiveWorkbook.Worksheets.Add();

            PivotTable_WorkSheet.Name = "Pivot Table";

            Excel.Range Paste_Range = PivotTable_WorkSheet.get_Range("A1");

            Paste_Range.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

        }
        public void Delete_Invalid_Rows(Excel.Worksheet RawData_Sheet, Excel.Range TimetoWrite_Range)
        {
            RawData_Sheet.AutoFilterMode = false;

            TimetoWrite_Range.AutoFilter(1, "=Invalid", Excel.XlAutoFilterOperator.xlAnd, System.Type.Missing, true);

            var xlFilteredRange = TimetoWrite_Range.Offset[1, 0].SpecialCells(Excel.XlCellType.xlCellTypeVisible, System.Type.Missing);

            xlFilteredRange.EntireRow.Delete(Excel.XlDirection.xlUp);

            RawData_Sheet.AutoFilterMode = false;

        }
        public Excel.Range Filter_By_Date(Excel.Worksheet RawData_Sheet, Excel.Range TestEnd_Range)
        {
            RawData_Sheet.AutoFilterMode = false;

            TestEnd_Range.AutoFilter(1, "<=" + Before_This_Date.Value.ToShortDateString(), Excel.XlAutoFilterOperator.xlAnd, ">=" + After_This_Date.Value.ToShortDateString(), true);

            dynamic allDataRange = RawData_Sheet.UsedRange.Offset[1,0];

            allDataRange.Sort(allDataRange.Columns[5], Excel.XlSortOrder.xlAscending);

            Excel.Range Filtered_Range = RawData_Sheet.UsedRange.SpecialCells(Excel.XlCellType.xlCellTypeVisible, System.Type.Missing);

            return Filtered_Range;
            
        }

        public void Add_Metric_Calculations(Excel.Worksheet RawData_Sheet)
        {
            Excel.Range Metric1 = RawData_Sheet.get_Range("K1", "K" + RawData_Sheet.UsedRange.Rows.Count);
            Excel.Range Metric2 = RawData_Sheet.get_Range("L1", "L" + RawData_Sheet.UsedRange.Rows.Count);

            Metric2.Formula = "=IF((H1-G1)>=0, H1-G1, \" Invalid\" )";

            Metric1.Formula = "=IF(AND(B1=\"A\" , (G1-E1)>=0, F1 <> \"May\", F1 <> \"Shah\"), G1-E1, \" Invalid\" )";

            Metric1.Cells[1] = "Time for Report";
            Metric2.Cells[1] = "Time for Check";
        }

        public void Delete_Empty_Rows(Excel.Worksheet RawData_Sheet, Excel.Range TestEnd_Range)
        {
            RawData_Sheet.AutoFilterMode = false;
            
            TestEnd_Range.AutoFilter(1, "=", Excel.XlAutoFilterOperator.xlAnd, System.Type.Missing, true);

            var xlFilteredRange = TestEnd_Range.Offset[1, 0].SpecialCells(Excel.XlCellType.xlCellTypeVisible, System.Type.Missing);

            xlFilteredRange.EntireRow.Delete(Excel.XlDirection.xlUp);

            RawData_Sheet.AutoFilterMode = false;
        }

        public void Delete_Unused_Columns(Excel.Worksheet RawData_Sheet)
        {
            Excel.Range Unused_Range = RawData_Sheet.get_Range("E1", "N" + RawData_Sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row);

            Unused_Range.Delete();

            Excel.Range Unused_Range_2 = RawData_Sheet.get_Range("K1", "AI" + RawData_Sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row);

            Unused_Range_2.Delete();
        }
    }
}
