using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SpreadCheck
{
	public partial class Form1 : Form
    {
        public int endcolumn = 0;
        public int foundLastRow = 0;
		public static bool rulesrunning = true;
        public static ColumnRules[] RuleList = new ColumnRules[200];

		DateTimeForm dateTimeForm = new DateTimeForm();
		ReportingOptions reportForm = new ReportingOptions();
		Progress progress = new Progress();

		public class Position
        {   public Position(int row, int col) { this.Row = row; this.Col = col; }
            public Position() { Row = 0; Col = 0; }
            public int Row { get; set; }
            public int Col { get; set; }
        }

        private Excel.Application xlApp = new Excel.Application();
        public static Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Range range;
        public static Reporter Report;
        object misvalue = System.Reflection.Missing.Value;

        public Excel.Worksheet XlWorkSheet { get => xlWorkSheet; set => xlWorkSheet = value; }
        public Excel.Range Range { get => range; set => range = value; }
        public Excel.Application XlApp { get => xlApp; set => xlApp = value; }

        public Form1(){InitializeComponent();}
        public int GetEndColumn() { return endcolumn; }

        private void Button1_Click(object sender, EventArgs e)
        {   RuleList[HeaderList.SelectedIndex].AllowedValuesArray = new List<string>();
            if (ItemCheckTextBox.TextLength > 0)
            {	AllowedItemsList.Items.Add(ItemCheckTextBox.Text);
                ItemCheckTextBox.Text = "";
                    foreach (string ListBoxItem in AllowedItemsList.Items)
                    {	RuleList[HeaderList.SelectedIndex].AllowedValuesArray.Add(ListBoxItem.ToString());		}
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {   if (AllowedItemsList.SelectedIndex > -1)
            {	AllowedItemsList.Items.Add(ItemCheckTextBox.Text);		}
            else
            {	MessageBox.Show("Item Not Selected");	}
        }
		/*  --- OPEN FILE ------*/
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {   int MaxCount = 100;
			RuleList = null;
			RuleList = new ColumnRules [100];
			HeaderList.Items.Clear();

            for (int b = 0; b < MaxCount; b++) //initialise
            { RuleList[b] = new ColumnRules(); }
	
            try
            {	openExcelFileDialog.ShowDialog();
                StatusLabel.Text = "Initialising Excel Engine...";
				xlWorkBook = xlApp.Workbooks.Open(Filename:openExcelFileDialog.FileName, Notify:true, UpdateLinks: updateLinksOnOpenToolStripMenuItem.Checked);            
            }	catch { return; }
			ColumnHeaderStart.Text = "1";
			RowHeaderStart.Text = "1";
			StatusLabel.Text = "Connecting to Worksheet...";
			xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
			StatusLabel.Text = "Detecting Used Range...";
			range = xlWorkSheet.UsedRange;

			StatusLabel.Text = "Adding Worksheets...";
			foreach (Excel.Worksheet  Worksheet in xlWorkBook.Worksheets)
            { SheetComboBox.Items.Add(Worksheet.Name);		}
            SheetComboBox.SelectedIndex = 0;


            StatusLabel.Text = "Detecting End Column...";
            for (int a = (Convert.ToInt32(ColumnHeaderStart.Text.Trim())); a < MaxCount; a++)
            {   if( (xlWorkSheet.Cells[Convert.ToInt32(RowHeaderStart.Text.Trim()),a].Value2) == null)
                { endcolumn =a; a = MaxCount;}
                else {
                    HeaderList.Items.Add((string)xlWorkSheet.Cells[Convert.ToInt32(ColumnHeaderStart.Text.Trim()),a].Value2.ToString());
                    RuleList[a].ColumnName = (string)xlWorkSheet.Cells[Convert.ToInt32(ColumnHeaderStart.Text.Trim()), a].Value2;		}
                          
            }

			if (endcolumn == 0) { MessageBox.Show("No Columns Found!"); }
			else
			{	reportingToolStripMenuItem.Enabled = true;
				EnabledCheckBox.Enabled = true;
				HeaderList.SelectedIndex = 0;
				ColumnnHeaderEnd.Text = endcolumn.ToString(); }
            StatusLabel.Text = "Detecting Last Row...";
            foundLastRow= xlWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;
            ColumnRules.SetLastDetectedRow(foundLastRow);
            LastRow.Text = ColumnRules.ReturnLastDetectedRow().ToString();
            Array.Resize(ref RuleList, endcolumn);

            StatusLabel.Text = "Creating Report Sheet...";
            colrowlabel.Text = "C:" + endcolumn.ToString() + "R:" + foundLastRow.ToString();
           
           Report = new Reporter(xlWorkBook, xlWorkSheet, endcolumn);
           
			this.Text = "SpreadChecker - " + openExcelFileDialog.SafeFileName;

			//Check for existing settings... 

			string fullpath = Path.GetDirectoryName(Application.ExecutablePath) + @"\" + Path.GetFileNameWithoutExtension(openExcelFileDialog.FileName) + ".dat";
		
			reportForm.RuleSetLocation.Text = fullpath;

			if (File.Exists(fullpath))
				{ LoadSettings(fullpath); }

			previewColumnDataToolStripMenuItem.Enabled = true;
			StatusLabel.Text = "  Idle...";
		}

        private void releaseObject(object obj)
        {   try
            {	System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;		}
            catch (Exception ex)
            {   obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());	}
            finally
            {	GC.Collect();	}
        }

        public void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {   try {
				xlWorkBook.Close(true, openExcelFileDialog.FileName, null);
				xlApp.Quit();
				releaseObject(xlWorkSheet);
				releaseObject(xlWorkBook);
				releaseObject(xlApp);
			} catch { MessageBox.Show("Save Cancelled"); }
		
        }

        private void CheckedChanged(object sender, EventArgs e)
        {   if (sender == WarnIfEmpty) { RuleList[HeaderList.SelectedIndex].IsEmpty = WarnIfEmpty.Checked; }
            if (sender == WarnIfContainsSpace) { RuleList[HeaderList.SelectedIndex].ContainsSpaces = WarnIfContainsSpace.Checked; }
            if (sender == IfNumbersCheckBox) { RuleList[HeaderList.SelectedIndex].ContainsNumber = IfNumbersCheckBox.Checked; }
            if (sender == IfLettersCheckBox) { RuleList[HeaderList.SelectedIndex].ContainsLetters = IfLettersCheckBox.Checked; }
            if (sender == IfNonAlphaCheckBox) { RuleList[HeaderList.SelectedIndex].ContainsNonAlpha = IfNonAlphaCheckBox.Checked; }
            if (sender == MustBeginWithCheckbox) { RuleList[HeaderList.SelectedIndex].CheckMustBeginWith = MustBeginWithCheckbox.Checked; }
            if (sender == MustEndWithCheckBox) { RuleList[HeaderList.SelectedIndex].CheckMustEndWith = MustEndWithCheckBox.Checked; }
            if (sender == LengthEnabledCheckbox) { RuleList[HeaderList.SelectedIndex].CheckLength = LengthEnabledCheckbox.Checked; }
            if (sender == AllowValuesCheckbox) { RuleList[HeaderList.SelectedIndex].AllowValuesEnabled = AllowValuesCheckbox.Checked; }
            if (sender == EnabledCheckBox) { RuleList[HeaderList.SelectedIndex].Enabled = EnabledCheckBox.Checked; }
            if (sender == LengthNumericUpDown) { RuleList[HeaderList.SelectedIndex].Length = (int)LengthNumericUpDown.Value; }
            if (sender == TrimCheckBox) { RuleList[HeaderList.SelectedIndex].TrimCell = TrimCheckBox.Checked; }
            if (sender == ReverseCheckBox) { RuleList[HeaderList.SelectedIndex].ReverseData = ReverseCheckBox.Checked; }
            if (sender == ChangeCaseCombo) { RuleList[HeaderList.SelectedIndex].ChangeCaseType = ChangeCaseCombo.SelectedIndex; }
            if (sender == BeginWithTextBox) { RuleList[HeaderList.SelectedIndex].MustBeginWith = BeginWithTextBox.Text; }
            if (sender == EndWithTextBox) { RuleList[HeaderList.SelectedIndex].MustEndWith = EndWithTextBox.Text; }
            if (sender == lessThanCheckBox) { RuleList[HeaderList.SelectedIndex].CheckLessThan = lessThanCheckBox.Checked; }
            if (sender == MoreThanCheckBox) { RuleList[HeaderList.SelectedIndex].CheckMoreThan = MoreThanCheckBox.Checked; }
            if (sender == MoreThanNumber) { RuleList[HeaderList.SelectedIndex].MoreThanValue = Convert.ToDouble(MoreThanNumber.Value); }
            if (sender == LessThanNumber) { RuleList[HeaderList.SelectedIndex].LessThanValue = Convert.ToDouble(LessThanNumber.Value); }
            if (sender == ChangeCaseCheckBox) { RuleList[HeaderList.SelectedIndex].ChangeCaseEnabled = ChangeCaseCheckBox.Checked; }
            if (sender == ColorCellCheckBox) { RuleList [HeaderList.SelectedIndex].BackColorEnabled = ColorCellCheckBox.Checked; }
			if (sender == TextRealignCheckbox) { RuleList [HeaderList.SelectedIndex].ChangeTextAlignment = TextRealignCheckbox.Checked; }
			if (sender == TextAlignmentList) { RuleList [HeaderList.SelectedIndex].TextAlignmentType = TextAlignmentList.SelectedIndex; }

            EndWithTextBox.Enabled = MustEndWithCheckBox.Checked;
            BeginWithTextBox.Enabled = MustBeginWithCheckbox.Checked;
            LengthNumericUpDown.Enabled = LengthEnabledCheckbox.Checked;

			ChangeCaseCombo.Enabled = ChangeCaseCheckBox.Checked;
			TextAlignmentList.Enabled = TextRealignCheckbox.Checked;
			RulesGroupBox.Enabled = EnabledCheckBox.Checked;
			MoreThanNumber.Enabled = MoreThanCheckBox.Checked;
			LessThanNumber.Enabled = lessThanCheckBox.Checked;

			int EnabledCount = 0;  //should probably have made this a method.
			for (int CountEnabled = 0; CountEnabled < endcolumn; CountEnabled++)
			{
				if (RuleList [CountEnabled].Enabled) { EnabledCount++; }
			}
				int CellCount = (EnabledCount * foundLastRow);
				StatLabel.Text = EnabledCount.ToString() + " RuleSets Enabled " + CellCount.ToString() + " Cells" ;

			if (EnabledCount > 0) RunButton.Enabled = true; else RunButton.Enabled = false;

		}

        private void HeaderList_SelectedIndexChanged(object sender, EventArgs e)
        {   WarnIfEmpty.Checked = RuleList[HeaderList.SelectedIndex].IsEmpty;
            WarnIfContainsSpace.Checked = RuleList[HeaderList.SelectedIndex].ContainsSpaces;
            IfNumbersCheckBox.Checked = RuleList[HeaderList.SelectedIndex].ContainsNumber;
            IfLettersCheckBox.Checked = RuleList[HeaderList.SelectedIndex].ContainsLetters;
            IfNonAlphaCheckBox.Checked = RuleList[HeaderList.SelectedIndex].ContainsNonAlpha;

            MustBeginWithCheckbox.Checked = RuleList[HeaderList.SelectedIndex].CheckMustBeginWith;
            BeginWithTextBox.Text = RuleList[HeaderList.SelectedIndex].MustBeginWith;

            MustEndWithCheckBox.Checked = RuleList[HeaderList.SelectedIndex].CheckMustEndWith;
            EndWithTextBox.Text = RuleList[HeaderList.SelectedIndex].MustEndWith;

            LengthEnabledCheckbox.Checked = RuleList[HeaderList.SelectedIndex].CheckLength;
            AllowValuesCheckbox.Checked = RuleList[HeaderList.SelectedIndex].AllowValuesEnabled;
            EnabledCheckBox.Checked = RuleList[HeaderList.SelectedIndex].Enabled;
            LengthNumericUpDown.Value = RuleList[HeaderList.SelectedIndex].Length;
            LengthNumericUpDown.Enabled = LengthEnabledCheckbox.Checked;
            TrimCheckBox.Checked = RuleList[HeaderList.SelectedIndex].TrimCell;
            ReverseCheckBox.Checked = RuleList[HeaderList.SelectedIndex].ReverseData;
            ChangeCaseCombo.SelectedIndex = RuleList[HeaderList.SelectedIndex].ChangeCaseType;

            lessThanCheckBox.Checked = RuleList[HeaderList.SelectedIndex].CheckLessThan;
            MoreThanCheckBox.Checked = RuleList[HeaderList.SelectedIndex].CheckMoreThan;

            MoreThanNumber.Value = Convert.ToDecimal(RuleList[HeaderList.SelectedIndex].MoreThanValue);
            LessThanNumber.Value = Convert.ToDecimal(RuleList[HeaderList.SelectedIndex].LessThanValue);

            ChangeCaseCheckBox.Checked = RuleList[HeaderList.SelectedIndex].ChangeCaseEnabled;
			TextRealignCheckbox.Checked = RuleList [HeaderList.SelectedIndex].ChangeTextAlignment;
			ColorCellCheckBox.Checked = RuleList [HeaderList.SelectedIndex].BackColorEnabled;
			TextRealignCheckbox.Checked = RuleList [HeaderList.SelectedIndex].ChangeTextAlignment;
			TextAlignmentList.SelectedIndex = RuleList [HeaderList.SelectedIndex].TextAlignmentType;
            AllowedItemsList.Items.Clear();

            if(RuleList[HeaderList.SelectedIndex].AllowedValuesArray != null)
            {   foreach (string stItem in RuleList[HeaderList.SelectedIndex].AllowedValuesArray)
                {   if (stItem !=null)
                    AllowedItemsList.Items.Add(stItem);		}
            }

			var ReturnCellType = xlWorkSheet.Cells [2, (HeaderList.SelectedIndex + 1)].Value;

			if (ReturnCellType != null)
			{	CellFormatLabel.Text = ReturnCellType.GetType().Name;
				NumberFormatLabel.Text = xlWorkSheet.Cells [2, (HeaderList.SelectedIndex + 1)].NumberFormat.ToString();		}
			else
			{	CellFormatLabel.Text = "NULL!";
				NumberFormatLabel.Text = "NULL!";	}

			ChangeCaseCombo.Enabled = ChangeCaseCheckBox.Checked;
			TextAlignmentList.Enabled = TextRealignCheckbox.Checked;
			RulesGroupBox.Enabled = RuleList [HeaderList.SelectedIndex].Enabled;

			int EnabledCount = 0;
		
			for (int CountEnabled = 0; CountEnabled < endcolumn; CountEnabled++)
			{	if (RuleList [CountEnabled].Enabled) { EnabledCount++; }
			}
		
			int CellCount = (EnabledCount * foundLastRow);
			StatLabel.Text = EnabledCount.ToString() + " RuleSets Enabled " + CellCount.ToString() + " Cells";

			if (EnabledCount > 0) RunButton.Enabled = true; else RunButton.Enabled = false;
		}

        private void ItemCheckTextBox_TextChanged(object sender, EventArgs e)
        {        }

         /*  ----------------RUN BUTTON!!! -------------  */

        private void RunButton_Click(object sender, EventArgs e)
        {	StatusLabel.Text = "Running...";
			Stopwatch sw = new Stopwatch();
			sw.Start();
			TimeSpan time = new TimeSpan();
			progress.Show();
			
			progress.RunProgress.Visible = true;

			RulesGroupBox.Enabled = false;
			HeaderList.Enabled = false;
			RunButton.Enabled = false;
			
            var pos = new Position();
            var Check = new RuleCheck(XlWorkSheet);
            var xlFunc = new XLFunctions();
            int EnabledCount = 0;
			progress.RunProgress.Minimum = 0;
         
            for (int CountEnabled = 0; CountEnabled < endcolumn; CountEnabled++)
            {  if (RuleList[CountEnabled].Enabled) { EnabledCount++; }
            }

			progress.RunProgress.Maximum = foundLastRow  * EnabledCount;
            Form1.Report.MakeHeaders(xlWorkSheet, Convert.ToInt32(ColumnHeaderStart.Text.Trim()),(Convert.ToInt32(ColumnHeaderStart.Text.Trim()) + endcolumn), Convert.ToInt32(RowHeaderStart.Text.Trim()),RuleList);

			for (int Row = Convert.ToInt32(RowHeaderStart.Text.Trim()) + 1; Row < ColumnRules.ReturnLastDetectedRow(); Row++) {
				for (int Col = Convert.ToInt32(ColumnHeaderStart.Text.Trim()); Col < endcolumn; Col++) {
					if (!rulesrunning) {
						progress.StopButton.Visible = false;
						
						return;
					}
					try {
						if (RuleList [Col - 1].Enabled)  //If a rule is enabled, begin processing
						{
							pos.Col = Col; pos.Row = Row;
							progress.RowCheckLabel.Text = "Row:" + Row.ToString();
							progress.ColumnCheckLabel.Text = "Column:" + Col.ToString();

							if (RuleList [Col - 1].IsEmpty)  //Check for empty cell
							{ Check.CheckifEmpty(xlWorkSheet.Cells [Row, Col].Value2, pos); }
							if (RuleList [Col - 1].CheckLength) { Check.CheckLength(Cell: xlWorkSheet.Cells [Row, Col].Value, pos: pos, Length: RuleList [Col - 1].Length); }
							if (RuleList [Col - 1].ContainsSpaces) { Check.CheckForSpaces(xlWorkSheet.Cells [Row, Col].Value, pos); }
							if (RuleList [Col - 1].ContainsNumber) { Check.CheckForNumbers(XlWorkSheet.Cells [Row, Col].Value, pos); }
							if (RuleList [Col - 1].AllowValuesEnabled) { Check.CheckifListEqual(xlWorkSheet.Cells [Row, Col].Value, pos, RuleList [Col - 1].AllowedValuesArray); }
							if (RuleList [Col - 1].ContainsNonAlpha) { Check.CheckForNonAlpha(xlWorkSheet.Cells [Row, Col].Value, pos); }
							if (RuleList [Col - 1].ContainsLetters) { Check.CheckForLetter(xlWorkSheet.Cells [Row, Col].Value, pos); }
							if (RuleList [Col - 1].CheckMustBeginWith) { Check.CheckBeginsWith(xlWorkSheet.Cells [Row, Col].Value, pos, RuleList [Col - 1].MustBeginWith); }
							if (RuleList [Col - 1].CheckMustEndWith) { Check.CheckEndsWith(xlWorkSheet.Cells [Row, Col].Value, pos, RuleList [Col - 1].MustEndWith); }
							if (RuleList [Col - 1].ChangeCaseEnabled) { xlWorkSheet.Cells [Row, Col].Value2 = xlFunc.ChangeCase(xlWorkSheet.Cells [Row, Col].Value, RuleList [Col - 1].ChangeCaseType); }
							if (RuleList [Col - 1].TrimCell) { xlWorkSheet.Cells [Row, Col].Value = xlFunc.TrimCell(xlWorkSheet.Cells [Row, Col].Value); }
							if (RuleList [Col - 1].ReverseData) { xlWorkSheet.Cells [Row, Col].Value = xlFunc.ReverseCell(xlWorkSheet.Cells [Row, Col].Value); }
							if (RuleList[Col - 1].CheckMoreThan) { Check.CheckNumber(XlWorkSheet.Cells [Row, Col].Value, pos, RuleList [Col - 1].MoreThanValue, true); }
							if (RuleList [Col - 1].CheckLessThan) { Check.CheckNumber(XlWorkSheet.Cells [Row, Col].Value, pos, RuleList [Col - 1].LessThanValue, false); }

							if (progress.RunProgress.Value < progress.RunProgress.Maximum)
								progress.RunProgress.Value++;
							progress.ErrorsFoundLabel.Text = "Errors Found:" + Report.GetErrorNumbers().ToString();
							time = TimeSpan.FromMilliseconds(sw.ElapsedMilliseconds);
							string answer = string.Format("{0:D2}h:{1:D2}m:{2:D2}s", time.Hours, time.Minutes, time.Seconds);
							progress.ElapsedLabel.Text = "Time Elapsed:" + answer;
							progress.FunctionsRunLabel.Text = "Functions Run:" + xlFunc.FunctionCallCount.ToString();

							Application.DoEvents();
						}
					} catch (Exception ex){ MessageBox.Show(ex.ToString()); }
				}
			}

			Form1.Report.CleanUpReport();
			progress.StopButton.Visible = false;
			
			StatusLabel.Text = "Complete!...";
        }

        private void Button2_Click_1(object sender, EventArgs e)
        {	try
			{	AllowedItemsList.Items.RemoveAt(AllowedItemsList.SelectedIndex);
				int b= 0;

				foreach (string Array in RuleList [HeaderList.SelectedIndex].AllowedValuesArray)
				{	RuleList [HeaderList.SelectedIndex].AllowedValuesArray [b] = "";
					b++;		}

				int a= 0;

				foreach (var ListBoxItem in AllowedItemsList.Items)
				{	RuleList [HeaderList.SelectedIndex].AllowedValuesArray [a] = AllowedItemsList.Items [a].ToString();
					a++;		}
			}	catch { };
        }

        private void RowHeaderStart_Click(object sender, EventArgs e)
        {       }

        private void button3_Click(object sender, EventArgs e)
        { dateTimeForm.ShowDialog(); }

        private void ChangeCaseCombo_ControlAdded(object sender, ControlEventArgs e)
        { ChangeCaseCombo.SelectedIndex = 1;}

		private void reportingToolStripMenuItem_Click(object sender, EventArgs e)
		{ reportForm.ShowDialog();}

		private void RulesGroupBox_Enter(object sender, EventArgs e)
		{		}

		private void Form1_Load(object sender, EventArgs e)
		{	foreach( string strNumberFormat in Enum.GetNames(typeof(Excel.XlListDataType)))
			{	NumberFormatCombo.Items.Add(strNumberFormat);	}
		}

		private void NumberFormatCombo_SelectedIndexChanged(object sender, EventArgs e)
		{		}

		private void enableAllRuleSetsToolStripMenuItem_Click(object sender, EventArgs e)
		{	int i = 0;
			foreach (string  str in HeaderList.Items)
			{	RuleList [i].Enabled = true; i++;}
			HeaderList_SelectedIndexChanged(sender,e);
		}

		private void disableAllRuleSetsToolStripMenuItem_Click(object sender, EventArgs e)
		{	int i = 0;
			foreach (string str in HeaderList.Items)
			{	RuleList [i].Enabled = false;	i++;}
			HeaderList_SelectedIndexChanged(sender, e);
		}

		private void exitToolStripMenuItem_Click(object sender, EventArgs e)
		{	Form1_FormClosing(sender, null);	}

		private void SaveExitButton_Click(object sender, EventArgs e)
		{	Form1_FormClosing(sender, null);	}

		public void StopButton_Click(object sender, EventArgs e)
		{ rulesrunning = false;	}
		
		public void SaveSettings(string path) /* --- save settings! ---*/
		{	try {
				string onlyFileName = System.IO.Path.GetFileNameWithoutExtension(path);
				onlyFileName = onlyFileName + ".dat";
				FileInfo f = new FileInfo(path);
				using (BinaryWriter bw = new BinaryWriter(f.OpenWrite())) {
					StatusLabel.Text = "Saving Column / Row Position Data.... ";
					bw.Write(ColumnHeaderStart.Text);
					bw.Write(RowHeaderStart.Text);
					StatusLabel.Text = "Saving Error Strings.... ";
					bw.Write((long)RuleCheck.err.BackCellColor);
					bw.Write(RuleCheck.err.EmptyError);
					bw.Write(RuleCheck.err.SpacesError);
					bw.Write(RuleCheck.err.NonAlpaError);
					bw.Write(RuleCheck.err.IfNumbersError);
					bw.Write(RuleCheck.err.IfLettersError);
					bw.Write(RuleCheck.err.MustBeginWithError);
					bw.Write(RuleCheck.err.MustEndWithError);
					bw.Write(RuleCheck.err.MustHaveLenghtError);
					bw.Write(RuleCheck.err.AllowedItemsListError);
					bw.Write(RuleCheck.err.IfMoreThanError);
					bw.Write(RuleCheck.err.IfLessThanError);
					bw.Write(RuleCheck.err.StringEqual);
					bw.Write(RuleCheck.err.CellNull);
					bw.Write(RuleCheck.err.UnexpectedDataTypeError);
					bw.Write(RuleCheck.err.BackCellColor.ToString());
					bw.Write(Report.IncludeHyperLink);
					bw.Write(Report.IncludeOriginalData);
					bw.Write(Report.IncludeEnabledData);
					bw.Write(Report.UseHeaderNames);
					bw.Write(Report.IncludePivotTable);
					bw.Write(Report.ReportNull);
					bw.Write(Report.ReportXData);
					bw.Write(endcolumn);

					for (int a = 0; a < endcolumn; a++) {
						StatusLabel.Text = "Saving.... " + RuleList [a].ColumnName;
						bw.Write(RuleList [a].IsEmpty);
						bw.Write(RuleList [a].ContainsSpaces);
						bw.Write(RuleList [a].ContainsNumber);
						bw.Write(RuleList [a].ContainsNonAlpha);
						bw.Write(RuleList [a].ContainsLetters);
						bw.Write(RuleList [a].CheckDateTime);
						//bw.Write((double) RuleList[a].LowerDate);
						//bw.Write(RuleList[a].UpperDate);
						bw.Write(RuleList [a].Enabled);
						bw.Write(RuleList [a].AllowValuesEnabled);
						bw.Write(RuleList [a].ColumnName);
						bw.Write(RuleList [a].CheckMustBeginWith);
						bw.Write(RuleList [a].MustBeginWith);
						bw.Write(RuleList [a].CheckMustEndWith);
						bw.Write(RuleList [a].MustEndWith);
						bw.Write(RuleList [a].CheckMoreThan);
						bw.Write(RuleList [a].CheckLessThan);
						bw.Write(RuleList [a].MoreThanValue);
						bw.Write(RuleList [a].LessThanValue);

						bw.Write(RuleList [a].CheckLength);
						bw.Write(RuleList [a].Length);
						bw.Write(RuleList [a].HeaderColumnPosition);
						bw.Write(RuleList [a].HeaderRowPosition);

						bw.Write(RuleList [a].TrimCell);
						bw.Write(RuleList [a].ReverseData);
						bw.Write(RuleList [a].ChangeCaseType);
						bw.Write(RuleList [a].ChangeCaseEnabled);
						bw.Write(RuleList [a].ChangeTextAlignment);
						bw.Write(RuleList [a].TextAlignmentType);
						bw.Write(RuleList [a].BackColorEnabled);
						bw.Write(RuleList [a].AllowedValuesArray.Count);
						if(RuleList [a].AllowedValuesArray.Count > 0) {
							foreach(string s in RuleList [a].AllowedValuesArray) {
								bw.Write(s);
							}
						}
					}
					StatusLabel.Text = "Idle.... ";
				}
			} catch { };
		}

		public void LoadSettings (string path)
		{	try {
				string onlyFileName = System.IO.Path.GetFileNameWithoutExtension(path);
				onlyFileName = onlyFileName + ".dat";
				FileInfo f = new FileInfo(path);
				using (BinaryReader br = new BinaryReader(f.OpenRead())) {
					ColumnHeaderStart.Text=br.ReadString();
					RowHeaderStart.Text = br.ReadString();
					StatusLabel.Text = "Loading Error Strings.... ";
					RuleCheck.err.BackCellColor = (Excel.XlRgbColor)br.ReadInt64();
					RuleCheck.err.EmptyError = br.ReadString();
					RuleCheck.err.SpacesError = br.ReadString();
					RuleCheck.err.NonAlpaError = br.ReadString();
					RuleCheck.err.IfNumbersError = br.ReadString();
					RuleCheck.err.IfLettersError = br.ReadString();
					RuleCheck.err.MustBeginWithError = br.ReadString();
					RuleCheck.err.MustEndWithError = br.ReadString();
					RuleCheck.err.MustHaveLenghtError = br.ReadString();
					RuleCheck.err.AllowedItemsListError = br.ReadString();
					RuleCheck.err.IfMoreThanError = br.ReadString();
					RuleCheck.err.IfLessThanError = br.ReadString();
					RuleCheck.err.StringEqual = br.ReadString();
					RuleCheck.err.CellNull = br.ReadString();
					RuleCheck.err.UnexpectedDataTypeError = br.ReadString();
					RuleCheck.err.BackCellColor = (Excel.XlRgbColor)Enum.Parse(typeof(Excel.XlRgbColor), br.ReadString());

					Report.IncludeHyperLink = br.ReadBoolean();
					Report.IncludeOriginalData = br.ReadBoolean();
					Report.IncludeEnabledData = br.ReadBoolean();
					Report.UseHeaderNames = br.ReadBoolean();
					Report.IncludePivotTable = br.ReadBoolean();
					Report.ReportNull = br.ReadBoolean();
					Report.ReportXData = br.ReadBoolean();
					endcolumn = br.ReadInt32();

					for (int a = 0; a < endcolumn; a++) {
						StatusLabel.Text = "Loading Column .... ";
						RuleList [a].IsEmpty = br.ReadBoolean();
						RuleList [a].ContainsSpaces = br.ReadBoolean();
						RuleList [a].ContainsNumber = br.ReadBoolean();
						RuleList [a].ContainsNonAlpha = br.ReadBoolean();
						RuleList [a].ContainsLetters = br.ReadBoolean();
						RuleList [a].CheckDateTime = br.ReadBoolean();
						//(double) RuleList[a].LowerDate
						//RuleList[a].UpperDate
						RuleList [a].Enabled = br.ReadBoolean();
						RuleList [a].AllowValuesEnabled = br.ReadBoolean();
						//RuleList[a].AllowedValuesArray
						RuleList [a].ColumnName = br.ReadString();
						RuleList [a].CheckMustBeginWith = br.ReadBoolean();
						RuleList [a].MustBeginWith = br.ReadString();
						RuleList [a].CheckMustEndWith = br.ReadBoolean();
						RuleList [a].MustEndWith = br.ReadString();
						RuleList [a].CheckMoreThan = br.ReadBoolean();
						RuleList [a].CheckLessThan = br.ReadBoolean();
						RuleList [a].MoreThanValue = br.ReadDouble();
						RuleList [a].LessThanValue = br.ReadDouble();

						RuleList [a].CheckLength = br.ReadBoolean();
						RuleList [a].Length = br.ReadInt32();
						RuleList [a].HeaderColumnPosition = br.ReadInt32();
						RuleList [a].HeaderRowPosition = br.ReadInt32();

						RuleList [a].TrimCell = br.ReadBoolean();
						RuleList [a].ReverseData = br.ReadBoolean();
						RuleList [a].ChangeCaseType = br.ReadInt32();
						RuleList [a].ChangeCaseEnabled = br.ReadBoolean();
						RuleList [a].ChangeTextAlignment = br.ReadBoolean();
						RuleList [a].TextAlignmentType = br.ReadInt32();
						RuleList [a].BackColorEnabled = br.ReadBoolean();
						int tempcount = br.ReadInt32();
						RuleList [a].AllowedValuesArray = new List<string>();
						for (a = 0; a == tempcount; a++) { RuleList [a].AllowedValuesArray.Add(br.ReadString()); }

					}
				}
			} catch (Exception Ex){ MessageBox.Show(@"Error Loading Settings.. Possible Mismatch:= "+ Ex.ToString()); }
				StatusLabel.Text = "Idle.... ";
			try { HeaderList.SelectedIndex = 0; } catch { MessageBox.Show("Unable to Select Column"); }
		}


		private void saveRuleSetsToolStripMenuItem_Click(object sender, EventArgs e)
		{	saveSettingsDialog.InitialDirectory = Path.GetDirectoryName(Application.ExecutablePath);
			StatusLabel.Text = "Saving RuleSets.... ";
			try { saveSettingsDialog.ShowDialog(); } catch { };
			SaveSettings(saveFileDialog.FileName);

			StatusLabel.Text = "Idle.... ";
		}

		private void loadRuleSetsToolStripMenuItem_Click(object sender, EventArgs e)
		{	StatusLabel.Text = "Loading RuleSets.... ";
			try { openSettingsDialog.ShowDialog(); } catch { }
			LoadSettings(openSettingsDialog.FileName);
				StatusLabel.Text = "Idle.... ";
		}

		private void previewColumnDataToolStripMenuItem_Click(object sender, EventArgs e)
		{
			PreviewData pForm = new PreviewData(xlWorkBook, xlWorkSheet, Convert.ToInt32(LastRow.Text),HeaderList.SelectedIndex);
			
			pForm.ShowDialog();
		}

		private void updateLinksOnOpenToolStripMenuItem_Click(object sender, EventArgs e)
		{
			if (this.updateLinksOnOpenToolStripMenuItem.Checked == true)
				this.updateLinksOnOpenToolStripMenuItem.Checked = false; 
			else this.updateLinksOnOpenToolStripMenuItem.Checked = true;

		}
	}
}
