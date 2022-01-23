using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;





namespace SpreadCheck
{
    public  class  Reporter
    {   public Excel.Workbook xlWrkBk;
		Excel.Worksheet xlErrorSheet;
		Excel.Range xlSrcRange, xlDstRange;
		public bool IncludeHyperLink { get; set; }
		public bool IncludeOriginalData { get; set; }
		public bool IncludeEnabledData { get; set; }
		public bool IncludeFullRowOfData { get; set; }
		public bool UseHeaderNames { get; set; }
		public bool ReportNull { get; set; }
		public bool ReportXData { get; set; }

		public Excel.XlRgbColor CellBackColor { get; set;}
		protected  int ReportRowCounter = 2;
        private int EndColumn { get; set; }
		public int GetErrorNumbers() { return ReportRowCounter -2; }
        private Excel.Worksheet currentSheet;
		public bool IncludePivotTable { get; set;}

    public Reporter(Excel.Workbook xlWorkbook, Excel.Worksheet curSheet, int endPosition, List<string> headers)  //Constructor
		{	xlWrkBk = xlWorkbook;
            currentSheet = curSheet;
            EndColumn = endPosition;
            IncludeHyperLink = true;
			IncludeOriginalData = true;
			IncludeEnabledData = false;
			IncludeFullRowOfData = true;
			UseHeaderNames = true;
			CellBackColor = Excel.XlRgbColor.rgbDarkRed;
			IncludePivotTable = false;
			bool createErrorSheet = !headers.Contains("Errors");
			if (createErrorSheet)
			{
				xlErrorSheet = (Excel.Worksheet)xlWorkbook.Worksheets.Add(After:xlWorkbook.Sheets[xlWorkbook.Sheets.Count]);
				try
				{
					xlErrorSheet.Name = "Errors";
				}
				catch {
					MessageBox.Show(@"Unable to create Error Sheet"); 
				}
			}
			else
			{
				var result = MessageBox.Show(
					@"Clear Error Sheet", 
					@"Errors Sheet already exists, would you like to clear the sheet?", 
					MessageBoxButtons.YesNo, 
					MessageBoxIcon.Question);
				xlErrorSheet = (Excel.Worksheet) xlWorkbook.Worksheets["Errors"];
				if (result == DialogResult.Yes)
				{
					
					xlErrorSheet.Cells.ClearContents();
				}
			}
		}

        public void MakeHeaders(Excel.Worksheet xlWorksheet, int HeaderStart, int HeaderEnd, int HeaderRow, ColumnRules [] rules)
        {	xlErrorSheet.Cells[1, 1] = "COLUMN";
            xlErrorSheet.Cells[1, 2] = "ROW";
            xlErrorSheet.Cells[1, 3] = "Error";
            xlErrorSheet.Cells[1, 4] = "Expected";
            xlErrorSheet.Cells[1, 5] = "Recieved";
			xlErrorSheet.Rows[1].Interior.Color = Excel.XlRgbColor.rgbDarkCyan;

			if (Form1.Report.IncludeOriginalData)
			{	for (int a = 6; a < (EndColumn + 5); a++)
				{	xlErrorSheet.Cells [1, a] = rules [a - 5].ColumnName;	}
			}
			/*
            Excel.Range source, Desintation;
            string rString = null;
            rString = HeaderStart.ToString() + ":" + HeaderEnd.ToString();
            
           //source = xlWorksheet.Range[1,1,1,8]*/

        }

        public void CreatePivots( Excel.Range DataRange)
        {	Excel.Worksheet pSheet = xlWrkBk.Worksheets.Add();
			pSheet.Name = "Pivot Tables Errors";
			Excel.Range pRange = pSheet.Range["A1","A1"];

			//Excel.PivotCache pivotCache = (Excel.PivotCache)xlWrkBk.PivotCaches().Add(Excel.XlPivotTableSourceType.xlDatabase, DataRange);

			Excel.PivotCache pivotCache = (Excel.PivotCache)Form1.XlWorkBook.PivotCaches().Add(Excel.XlPivotTableSourceType.xlDatabase, DataRange);
			var pivotTables = (Excel.PivotTables)xlErrorSheet.PivotTables();
			var pivotTable = pivotTables.Add(pivotCache, pRange, "PivotTable1");
        }

        public void Add(int col, int row, string Error, object Expected, object Recieved, Excel.Worksheet sheet)
        {	if (Form1.RuleList [col - 1].BackColorEnabled)
				currentSheet.Cells [row, col].Interior.Color = Form1.Report.CellBackColor;

			if (UseHeaderNames)
				xlErrorSheet.Cells [ReportRowCounter, 1].Value = Form1.RuleList [col].ColumnName;
			else
				xlErrorSheet.Cells [ReportRowCounter, 1].Value = col;

			xlErrorSheet.Cells[ReportRowCounter, 2].Value = row;
            xlErrorSheet.Cells[ReportRowCounter, 3].Value = Error;
            xlErrorSheet.Cells[ReportRowCounter, 4].Value = Expected.ToString();
            xlErrorSheet.Cells[ReportRowCounter, 5].Value = Recieved.ToString();
		
            xlErrorSheet.Cells[ReportRowCounter, 6].Value = "\"= HYPERLINK(\"[Book1]Data!D10\"; \"name of hyperlink\")\"";     ;
            string hyperlinkTargetAddress = "Data!A1";

            int b = 6;
            col = 1;

			if (Form1.Report.IncludeOriginalData) //Include orginal data
			{	if (Form1.Report.IncludeFullRowOfData)
				{/*
					Excel.Range srange = sheet.Range [row, EndColumn].EntireRow;
					Excel.Range drange = xlErrorSheet.get_Range(ReportRowCounter, EndColumn + 6);
					srange.Copy();
					drange.PasteSpecial();
					*/
					for (; b < (EndColumn + 6); b++) // Copy out full column *slow*
					{
						xlErrorSheet.Cells [ReportRowCounter, b].Value = sheet.Cells [row, col].Value;
						col++;
					}
				}
				else // copy out enabled columsn only
				{	for (; b < (EndColumn + 6); b++) // Copy out Enabled Columns only
					{	if (Form1.RuleList [b - 6].Enabled)
						{	xlErrorSheet.Cells [ReportRowCounter, b].Value = sheet.Cells [row, col].Value;			}
						col++;
					}
				}
			}

			ReportRowCounter++;
        } 

		public void CleanUpReport()	{
			xlErrorSheet.Columns.AutoFit();
			xlErrorSheet.Rows.AutoFit();
			xlDstRange = xlErrorSheet.Range["A1","E5"];
			if(Form1.Report.IncludePivotTable)
			CreatePivots(xlDstRange);
		}

        public class ErrorMessages{
			public Excel.XlRgbColor BackCellColor { get; set; }
			public string EmptyError { get; set; }
            public string SpacesError { get; set; }
            public string NonAlpaError { get; set; }
            public string IfNumbersError { get; set; }
            public string IfLettersError { get; set; }
			public string MustBeginWithError { get; set; }
            public string MustEndWithError { get; set; }
            public string MustHaveLenghtError { get; set; }
            public string AllowedItemsListError { get; set; }
            public string IfMoreThanError { get; set; }
            public string IfLessThanError { get; set; }
			public string StringEqual { get; set; }
			public string CellNull { get; set; }
			public string UnexpectedDataTypeError { get; set; }

			public ErrorMessages(){
				EmptyError = @"Cell Empty!";
				SpacesError = @"Cell Contains a Space";
				NonAlpaError = @"Cell Contains Non-AlphaNumerics";
				IfNumbersError = @"Cell Contains Numbers";
				IfLettersError = @"Cell Contains Letters";
				MustBeginWithError = @"Cell Failed to Begin with";
				MustEndWithError = @"Cell Failed to end with";
				MustHaveLenghtError = @"Cell Failed to meet Length Criterior";
				AllowedItemsListError = "Cell Contents not Not Allowed";
				IfMoreThanError = "Cells Value Higher than Allowed";
				IfLessThanError = "Cell Value Lower than Allowed";
				StringEqual = "String is Not Equal";
				CellNull = "Cell NULL";
				UnexpectedDataTypeError = "Recieved Incorrect DataType for Check";
			}
        }


    }
}
