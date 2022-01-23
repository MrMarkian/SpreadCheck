using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace SpreadCheck
{
	public partial class Form1 : Form
    {
	    private int _endColumn;
	    private int _foundLastRow;
		public static bool RulesRunning = true;
        public static ColumnRules[] RuleList = new ColumnRules[200];

        private readonly DateTimeForm _dateTimeForm = new DateTimeForm();
        private readonly ReportingOptions _reportForm = new ReportingOptions();
        private readonly Progress _progress = new Progress();

		public class Position
        {   public Position(int row, int col) { Row = row; Col = col; }
            public Position() { Row = 0; Col = 0; }
            public int Row { get; set; }
            public int Col { get; set; }
        }

        private Application _xlApp = new Application();
        public static Workbook XlWorkBook;
        private Worksheet _xlWorkSheet;
        // private Range _range;
        public static Reporter Report;
        public object MissingValue { get; } = Missing.Value;

        public Worksheet XlWorkSheet { get => _xlWorkSheet; set => _xlWorkSheet = value; }
        // public Range Range { get => _range; set => _range = value; }
        public Application XlApp { get => _xlApp; set => _xlApp = value; }

        public Form1(){InitializeComponent();}
        public int GetEndColumn() { return _endColumn; }

        private void AddItemButton_Click(object sender, EventArgs e)
        {   RuleList[HeaderList.SelectedIndex].AllowedValuesArray = new List<string>();
            if (ItemCheckTextBox.TextLength > 0)
            {	AllowedItemsList.Items.Add(ItemCheckTextBox.Text);
                ItemCheckTextBox.Text = "";
                    foreach (string listBoxItem in AllowedItemsList.Items)
                    {	RuleList[HeaderList.SelectedIndex].AllowedValuesArray.Add(listBoxItem);		}
            }
        }

        // private void Button2_Click(object sender, EventArgs e)
        // {   if (AllowedItemsList.SelectedIndex > -1)
        //     {	AllowedItemsList.Items.Add(ItemCheckTextBox.Text);		}
        //     else
        //     {	MessageBox.Show("Item Not Selected");	}
        // }
		/*  --- OPEN FILE ------*/
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {   int MaxCount = 100;
			RuleList = null;
			RuleList = new ColumnRules [100];
			HeaderList.Items.Clear();

            for (int b = 0; b < MaxCount; b++) //initialise
            { RuleList[b] = new ColumnRules(); }

            if (TryOpenExcel()) return;
            SetRowColumnText();
            FindWorksheets();

            WorkOutWorkBookRange(MaxCount);
            Array.Resize(ref RuleList, _endColumn);

            StatusLabel.Text = @"Creating Report Sheet...";
            colrowlabel.Text = $@"C:{_endColumn.ToString()}R:{_foundLastRow.ToString()}";

            //TODO Check if Error Sheet Exists
            Report = new Reporter(XlWorkBook, _xlWorkSheet, _endColumn);
           
			Text = $@"SpreadChecker - {openExcelFileDialog.SafeFileName}";

			//Check for existing settings... 

			string fullPath =
				$@"{Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath)}\{Path.GetFileNameWithoutExtension(openExcelFileDialog.FileName)}.dat";
		
			_reportForm.RuleSetLocation.Text = fullPath;

			if (File.Exists(fullPath))
			{ LoadSettings(fullPath); }

			previewColumnDataToolStripMenuItem.Enabled = true;
			StatusLabel.Text = @"  Idle...";
		}

        private void WorkOutWorkBookRange(int maxCount)
        {
	        // StatusLabel.Text = @"Detecting Used Range...";
	        // not sure why this is needed at this point doesn't seam to be used
	        // _range = xlWorkSheet.UsedRange;

	        //TODO Might be worth having this auto detect using used range?
	        //TODO Test other functions
	        StatusLabel.Text = @"Detecting End Column...";
	        for (int a = (Convert.ToInt32(ColumnHeaderStart.Text.Trim())); a < maxCount; a++)
	        {
		        if ((_xlWorkSheet.Cells[Convert.ToInt32(RowHeaderStart.Text.Trim()), a].Value2) == null)
		        {
			        _endColumn = a;
			        a = maxCount;
		        }
		        else
		        {
			        HeaderList.Items.Add((string) _xlWorkSheet.Cells[Convert.ToInt32(ColumnHeaderStart.Text.Trim()), a].Value2
				        .ToString());
			        RuleList[a].ColumnName =
				        (string) _xlWorkSheet.Cells[Convert.ToInt32(ColumnHeaderStart.Text.Trim()), a].Value2;
		        }
	        }

	        if (_endColumn == 0)
	        {
		        MessageBox.Show(@"No Columns Found!");
	        }
	        else
	        {
		        reportingToolStripMenuItem.Enabled = true;
		        EnabledCheckBox.Enabled = true;
		        try
		        {
			        HeaderList.SelectedIndex = 0;
		        }
		        catch
		        {
			        MessageBox.Show(@"Loading Failed");
		        }

		        ColumnnHeaderEnd.Text = _endColumn.ToString();
	        }

	        StatusLabel.Text = @"Detecting Last Row...";
	        _foundLastRow = _xlWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
	        ColumnRules.SetLastDetectedRow(_foundLastRow);
	        LastRow.Text = ColumnRules.ReturnLastDetectedRow().ToString();
        }

        private void FindWorksheets()
        {
	        StatusLabel.Text = @"Connecting to Worksheet...";
	        //TODO only uses first tab needs tab option
	        _xlWorkSheet = (Worksheet) XlWorkBook.Worksheets.Item[1];

	        //TODO nothing using these worksheets, would be good to choose
	        StatusLabel.Text = @"Adding Worksheets...";
	        foreach (Worksheet worksheet in XlWorkBook.Worksheets)
	        {
		        SheetComboBox.Items.Add(worksheet.Name);
	        }

	        SheetComboBox.SelectedIndex = 0;
        }

        private void SetRowColumnText()
        {
	        //TODO Review sets up default row and column text box to 1
	        // why not use the use range for this? looks like this should only be set if no value is already in there
	        ColumnHeaderStart.Text = @"1";
	        RowHeaderStart.Text = @"1";
        }

        private bool TryOpenExcel()
        {
	        try
	        {
		        openExcelFileDialog.ShowDialog();
		        StatusLabel.Text = @"Initialising Excel Engine...";
		        XlWorkBook = _xlApp.Workbooks.Open(Filename: openExcelFileDialog.FileName, Notify: true,
			        UpdateLinks: updateLinksOnOpenToolStripMenuItem.Checked);
	        }
	        catch (Exception exception)
	        {
		        MessageBox.Show(exception.Message);
		        return true;
	        }

	        return false;
        }

        private void releaseObject(object obj)
        {   try
            {	Marshal.ReleaseComObject(obj);
                obj = null;		}
            catch (Exception ex)
            {   obj = null;
                MessageBox.Show($@"Unable to release the Object {ex.Message}");	}
            finally
            {	GC.Collect();	}
        }

        public void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {   try {
				XlWorkBook.Close(true, openExcelFileDialog.FileName);
				_xlApp.Quit();
				releaseObject(_xlWorkSheet);
				releaseObject(XlWorkBook);
				releaseObject(_xlApp);
			} catch { MessageBox.Show(@"Save Cancelled"); }
		
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

			int enabledCount = 0;  //should probably have made this a method.
			for (int countEnabled = 0; countEnabled < _endColumn; countEnabled++)
			{
				if (RuleList [countEnabled].Enabled) { enabledCount++; }
			}
			int cellCount = (enabledCount * _foundLastRow);
			StatLabel.Text = $@"{enabledCount.ToString()} RuleSets Enabled {cellCount.ToString()} Cells" ;

			if (enabledCount > 0) RunButton.Enabled = true; else RunButton.Enabled = false;

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

			var returnCellType = _xlWorkSheet.Cells [2, (HeaderList.SelectedIndex + 1)].Value;

			if (returnCellType != null)
			{	CellFormatLabel.Text = returnCellType.GetType().Name;
				NumberFormatLabel.Text = _xlWorkSheet.Cells [2, (HeaderList.SelectedIndex + 1)].NumberFormat.ToString();		}
			else
			{	CellFormatLabel.Text = @"NULL!";
				NumberFormatLabel.Text = @"NULL!";	}

			ChangeCaseCombo.Enabled = ChangeCaseCheckBox.Checked;
			TextAlignmentList.Enabled = TextRealignCheckbox.Checked;
			RulesGroupBox.Enabled = RuleList [HeaderList.SelectedIndex].Enabled;

			int enabledCount = 0;
		
			for (int countEnabled = 0; countEnabled < _endColumn; countEnabled++)
			{	if (RuleList [countEnabled].Enabled) { enabledCount++; }
			}
		
			int cellCount = (enabledCount * _foundLastRow);
			StatLabel.Text = $@"{enabledCount.ToString()} RuleSets Enabled {cellCount.ToString()} Cells";

			if (enabledCount > 0) RunButton.Enabled = true; else RunButton.Enabled = false;
		}

        private void ItemCheckTextBox_TextChanged(object sender, EventArgs e)
        {        }

         /*  ----------------RUN BUTTON!!! -------------  */

        private void RunButton_Click(object sender, EventArgs e)
        {	
	        //TODO Work out why theres no errors showing in the test worksheet for blanks
	        //TODO Last Column not working
	        StatusLabel.Text = @"Running...";
	        
			Stopwatch stopwatch = new Stopwatch();
			stopwatch.Start();

			DisableOptions();

			// ??
			Position cellPosition = new Position();
            RuleCheck check = new RuleCheck(XlWorkSheet);
            XLFunctions xlFunc = new XLFunctions();
            
            int enabledRuleListCount = 0;
            for (int ruleListIndex = 0; ruleListIndex < _endColumn; ruleListIndex++)
            {  if (RuleList[ruleListIndex].Enabled) { enabledRuleListCount++; }
            }

            SetUpProgressBar(enabledRuleListCount);
            
			// Converting user input string to number
			int columnStart = Convert.ToInt32(ColumnHeaderStart.Text.Trim());
			int rowStart = Convert.ToInt32(RowHeaderStart.Text.Trim());
			int columnEnd = columnStart + _endColumn;
            Report.MakeHeaders(_xlWorkSheet, columnStart, columnEnd, rowStart, RuleList);
            
			for (int row = rowStart + 1; row < _foundLastRow; row++) {
				for (int column = columnStart; column < _endColumn; column++)
				{
					if (IsStopped()) return;
					try {
						if (RuleList [column - 1].Enabled)  //If a rule is enabled, begin processing
						{	cellPosition.Col = column; cellPosition.Row = row;

							check.ApplyRulesToCell( cellPosition, xlFunc, _xlWorkSheet.Cells[row, column], RuleList[column -1]);

							UpdateProgressInfo(row, column, stopwatch, xlFunc);
						}
					} catch (Exception ex){ MessageBox.Show(ex.ToString()); }
				}
			}

			Report.CleanUpReport();
			_progress.StopButton.Visible = false;
			
			StatusLabel.Text = @"Complete!...";
        }

        

        private void UpdateProgressInfo(int row, int column, Stopwatch stopwatch, XLFunctions xlFunc)
        {
	        _progress.RowCheckLabel.Text = @"Row:" + row;
	        _progress.ColumnCheckLabel.Text = @"Column:" + column;
	        if (_progress.RunProgress.Value < _progress.RunProgress.Maximum)
		        _progress.RunProgress.Value++;
	        _progress.ErrorsFoundLabel.Text = @"Errors Found:" + Report.GetErrorNumbers();
	        var time = TimeSpan.FromMilliseconds(stopwatch.ElapsedMilliseconds);
	        string answer = $"{time.Hours:D2}h:{time.Minutes:D2}m:{time.Seconds:D2}s";
	        _progress.ElapsedLabel.Text = $@"Time Elapsed:{answer}";
	        _progress.FunctionsRunLabel.Text = $@"Functions Run:{xlFunc.FunctionCallCount.ToString()}";
	        System.Windows.Forms.Application.DoEvents();
        }

        private bool IsStopped()
        {
	        if (!RulesRunning)
	        {
		        _progress.StopButton.Visible = false;
		        return true;
	        }

	        return false;
        }

        private void DisableOptions()
        {
	        RulesGroupBox.Enabled = false;
	        HeaderList.Enabled = false;
	        RunButton.Enabled = false;
        }

        private void SetUpProgressBar(int enabledRuleListCount)
        {
	        //TODO Fix progress metre and make more obvious function complete
	        _progress.Show();
	        _progress.RunProgress.Visible = true;
	        _progress.RunProgress.Minimum = 0;
	        // Set progress by number of rows x columns with enabled rule lists
	        _progress.RunProgress.Maximum = _foundLastRow  * enabledRuleListCount;
        }

        private void RemoveItemButton_Click_1(object sender, EventArgs e)
        {	try
			{	AllowedItemsList.Items.RemoveAt(AllowedItemsList.SelectedIndex);
				for (int b= 0; b < RuleList [HeaderList.SelectedIndex].AllowedValuesArray.Count; b++)
				{	RuleList [HeaderList.SelectedIndex].AllowedValuesArray [b] = ""; }
				for (int a= 0; a < AllowedItemsList.Items.Count; a++)
				{	RuleList [HeaderList.SelectedIndex].AllowedValuesArray [a] = AllowedItemsList.Items [a].ToString(); }
			}
	        catch
	        { /* ignored */ }
        }

        private void RowHeaderStart_Click(object sender, EventArgs e)
        {       }

        private void DateTimeSettingsButton_Click(object sender, EventArgs e)
        { _dateTimeForm.ShowDialog(); }

        // private void ChangeCaseCombo_ControlAdded(object sender, ControlEventArgs e)
        // { ChangeCaseCombo.SelectedIndex = 1;}

		private void reportingToolStripMenuItem_Click(object sender, EventArgs e)
		{ _reportForm.ShowDialog();}

		private void RulesGroupBox_Enter(object sender, EventArgs e)
		{		}

		private void Form1_Load(object sender, EventArgs e)
		{	
			foreach( string strNumberFormat in Enum.GetNames(typeof(XlListDataType)))
			{	NumberFormatCombo.Items.Add(strNumberFormat);	}
		}

		private void NumberFormatCombo_SelectedIndexChanged(object sender, EventArgs e)
		{		}

		private void enableAllRuleSetsToolStripMenuItem_Click(object sender, EventArgs e)
		{	
			for (int i = 0; i < HeaderList.Items.Count; i++)
			{	RuleList [i].Enabled = true;}
			HeaderList_SelectedIndexChanged(sender,e);
		}

		private void disableAllRuleSetsToolStripMenuItem_Click(object sender, EventArgs e)
		{	
			for (int i = 0; i < HeaderList.Items.Count; i++)
			{	RuleList [i].Enabled = false;}
			HeaderList_SelectedIndexChanged(sender, e);
		}

		private void exitToolStripMenuItem_Click(object sender, EventArgs e)
		{	Form1_FormClosing(sender, null);	}

		// private void SaveExitButton_Click(object sender, EventArgs e)
		// {	Form1_FormClosing(sender, null);	}
		//
		// public void StopButton_Click(object sender, EventArgs e)
		// { RulesRunning = false;	}
		
		public void SaveSettings(string path) /* --- save settings! ---*/
		{	try {
				string onlyFileName = $"{Path.GetFileNameWithoutExtension(path)}.dat";
				FileInfo f = new FileInfo(path);
				using (BinaryWriter bw = new BinaryWriter(f.OpenWrite())) {
					StatusLabel.Text = @"Saving Column / Row Position Data.... ";
					bw.Write(ColumnHeaderStart.Text);
					bw.Write(RowHeaderStart.Text);
					StatusLabel.Text = @"Saving Error Strings.... ";
					bw.Write((long)RuleCheck.Err.BackCellColor);
					bw.Write(RuleCheck.Err.EmptyError);
					bw.Write(RuleCheck.Err.SpacesError);
					bw.Write(RuleCheck.Err.NonAlpaError);
					bw.Write(RuleCheck.Err.IfNumbersError);
					bw.Write(RuleCheck.Err.IfLettersError);
					bw.Write(RuleCheck.Err.MustBeginWithError);
					bw.Write(RuleCheck.Err.MustEndWithError);
					bw.Write(RuleCheck.Err.MustHaveLenghtError);
					bw.Write(RuleCheck.Err.AllowedItemsListError);
					bw.Write(RuleCheck.Err.IfMoreThanError);
					bw.Write(RuleCheck.Err.IfLessThanError);
					bw.Write(RuleCheck.Err.StringEqual);
					bw.Write(RuleCheck.Err.CellNull);
					bw.Write(RuleCheck.Err.UnexpectedDataTypeError);
					bw.Write(RuleCheck.Err.BackCellColor.ToString());
					bw.Write(Report.IncludeHyperLink);
					bw.Write(Report.IncludeOriginalData);
					bw.Write(Report.IncludeEnabledData);
					bw.Write(Report.UseHeaderNames);
					bw.Write(Report.IncludePivotTable);
					bw.Write(Report.ReportNull);
					bw.Write(Report.ReportXData);
					bw.Write(_endColumn);

					for (int a = 0; a < _endColumn; a++) {
						StatusLabel.Text = $@"Saving.... {RuleList [a].ColumnName}";
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
					StatusLabel.Text = @"Idle.... ";
				}
			}
			catch { /* ignored */ }
		}

		private void LoadSettings (string path)
		{	try {
				string onlyFileName = $"{Path.GetFileNameWithoutExtension(path)}.dat";
				FileInfo f = new FileInfo(path);
				using (BinaryReader br = new BinaryReader(f.OpenRead())) {
					ColumnHeaderStart.Text=br.ReadString();
					RowHeaderStart.Text = br.ReadString();
					StatusLabel.Text = @"Loading Error Strings.... ";
					RuleCheck.Err.BackCellColor = (XlRgbColor)br.ReadInt64();
					RuleCheck.Err.EmptyError = br.ReadString();
					RuleCheck.Err.SpacesError = br.ReadString();
					RuleCheck.Err.NonAlpaError = br.ReadString();
					RuleCheck.Err.IfNumbersError = br.ReadString();
					RuleCheck.Err.IfLettersError = br.ReadString();
					RuleCheck.Err.MustBeginWithError = br.ReadString();
					RuleCheck.Err.MustEndWithError = br.ReadString();
					RuleCheck.Err.MustHaveLenghtError = br.ReadString();
					RuleCheck.Err.AllowedItemsListError = br.ReadString();
					RuleCheck.Err.IfMoreThanError = br.ReadString();
					RuleCheck.Err.IfLessThanError = br.ReadString();
					RuleCheck.Err.StringEqual = br.ReadString();
					RuleCheck.Err.CellNull = br.ReadString();
					RuleCheck.Err.UnexpectedDataTypeError = br.ReadString();
					RuleCheck.Err.BackCellColor = (XlRgbColor)Enum.Parse(typeof(XlRgbColor), br.ReadString());

					Report.IncludeHyperLink = br.ReadBoolean();
					Report.IncludeOriginalData = br.ReadBoolean();
					Report.IncludeEnabledData = br.ReadBoolean();
					Report.UseHeaderNames = br.ReadBoolean();
					Report.IncludePivotTable = br.ReadBoolean();
					Report.ReportNull = br.ReadBoolean();
					Report.ReportXData = br.ReadBoolean();
					_endColumn = br.ReadInt32();

					for (int a = 0; a < _endColumn; a++) {
						StatusLabel.Text = @"Loading Column .... ";
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
						int tempCount = br.ReadInt32();
						RuleList [a].AllowedValuesArray = new List<string>();
						for (a = 0; a == tempCount; a++) { RuleList [a].AllowedValuesArray.Add(br.ReadString()); }

					}
				}
			} catch (Exception ex){ MessageBox.Show($@"Error Loading Settings.. Possible Mismatch:= {ex.Message}");}
			StatusLabel.Text = @"Idle.... ";
			try { HeaderList.SelectedIndex = 0; } catch { MessageBox.Show(@"Unable to Select Column"); }
		}


		private void saveRuleSetsToolStripMenuItem_Click(object sender, EventArgs e)
		{	saveSettingsDialog.InitialDirectory = Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
			StatusLabel.Text = @"Saving RuleSets.... ";
			try { saveSettingsDialog.ShowDialog(); }
			catch
			{ /* ignored */}
			SaveSettings(saveFileDialog.FileName);

			StatusLabel.Text = @"Idle.... ";
		}

		private void loadRuleSetsToolStripMenuItem_Click(object sender, EventArgs e)
		{	StatusLabel.Text = @"Loading RuleSets.... ";
			try { openSettingsDialog.ShowDialog(); }
			catch { /* ignored */ }

			LoadSettings(openSettingsDialog.FileName);
				StatusLabel.Text = @"Idle.... ";
		}

		private void previewColumnDataToolStripMenuItem_Click(object sender, EventArgs e)
		{
			PreviewData pForm = new PreviewData(XlWorkBook, Convert.ToInt32(LastRow.Text),HeaderList.SelectedIndex);
			
			pForm.ShowDialog();
		}

		private void updateLinksOnOpenToolStripMenuItem_Click(object sender, EventArgs e)
		{
			if (updateLinksOnOpenToolStripMenuItem.Checked)
				updateLinksOnOpenToolStripMenuItem.Checked = false; 
			else updateLinksOnOpenToolStripMenuItem.Checked = true;

		}
	}
}
