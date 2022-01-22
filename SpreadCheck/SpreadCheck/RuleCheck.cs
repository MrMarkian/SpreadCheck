using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;


namespace SpreadCheck
{
    class RuleCheck
    {
        private readonly Excel.Worksheet _currentSheet;
        public RuleCheck(Excel.Worksheet sheet) { _currentSheet = sheet; }
      
		public static readonly Reporter.ErrorMessages Err = new Reporter.ErrorMessages();
	
		public void ApplyRulesToCell(Form1.Position cellPosition, XLFunctions xlFunc, Range range, ColumnRules ruleList)
        {
	        //TODO Evaluate if functions should be bool or just change to method
	        if (ruleList.IsEmpty) CheckIfEmpty(range.Value2, cellPosition);
	        if (ruleList.CheckLength) CheckLength(cell: range.Value, pos: cellPosition, length: ruleList.Length);
	        if (ruleList.ContainsSpaces) CheckForSpaces(range.Value, cellPosition);
	        if (ruleList.ContainsNumber) CheckForNumbers(range.Value, cellPosition);
	        if (ruleList.AllowValuesEnabled) CheckIfListEqual(range.Value, cellPosition, ruleList.AllowedValuesArray);
	        if (ruleList.ContainsNonAlpha) CheckForNonAlpha(range.Value, cellPosition);
	        if (ruleList.ContainsLetters) CheckForLetter(range.Value, cellPosition);
	        if (ruleList.CheckMustBeginWith) CheckBeginsWith(range.Value, cellPosition, ruleList.MustBeginWith);
	        if (ruleList.CheckMustEndWith) CheckEndsWith(range.Value, cellPosition, ruleList.MustEndWith);
	        if (ruleList.ChangeCaseEnabled) range.Value2 = xlFunc.ChangeCase(range.Value, ruleList.ChangeCaseType);
	        if (ruleList.TrimCell) range.Value = xlFunc.TrimCell(range.Value);
	        if (ruleList.ReverseData) range.Value = xlFunc.ReverseCell(range.Value);
	        if (ruleList.CheckMoreThan) CheckNumber(range.Value, cellPosition, ruleList.MoreThanValue, true);
	        if (ruleList.CheckLessThan) CheckNumber(range.Value, cellPosition, ruleList.LessThanValue, false);
        }

		private bool CheckIfEmpty(object cell, Form1.Position pos) // complete
        {	switch (cell) {
				case string stringValue:
					if (String.IsNullOrWhiteSpace(stringValue))
						Form1.Report.Add(
							pos.Col, pos.Row, Err.EmptyError, "AnyThing", "WhiteSpace", _currentSheet);
					break;
				case double doubleValue:
					if(doubleValue==0.00)
					{
						Form1.Report.Add(
							pos.Col, pos.Row, $"{Err.EmptyError} -Double", 1, 0, _currentSheet);
						break;
					}
					else
					{
						break;
					}

				case null:
					if (Form1.Report.ReportNull)
						Form1.Report.Add(
							pos.Col, pos.Row, Err.CellNull, 1, 0, _currentSheet);
					break;
				default:
					if(Form1.Report.ReportXData)
						Form1.Report.Add(
							pos.Col, pos.Row, Err.UnexpectedDataTypeError, "String", 
							cell.GetType().ToString(), _currentSheet);
					break;
			}
			return false;
        }

		private bool CheckForNumbers (object cell, Form1.Position pos) // complete
        {	switch (cell) {
				case string stringValue:
					if (stringValue.Any(char.IsNumber)) 
						Form1.Report.Add(
							pos.Col, pos.Row, $"{Err.IfNumbersError} -String", "String", 
							stringValue, _currentSheet);
					return true;
				case null:
				if (Form1.Report.ReportNull)
					Form1.Report.Add(pos.Col, pos.Row, Err.CellNull, "A-Z", "NULL", _currentSheet);
				break;

				default:
				if (Form1.Report.ReportXData)
					Form1.Report.Add(
						pos.Col, pos.Row, Err.UnexpectedDataTypeError, "A-Z", 
						cell.GetType().ToString(), _currentSheet);
				break;
			}
			return false;
        }

        private bool CheckForSpaces (object cell, Form1.Position pos)
		{	switch (cell) {
				case string s:
					if (s.Any(Char.IsWhiteSpace))
						Form1.Report.Add(
							pos.Col, pos.Row, Err.SpacesError, "No Spaces", s, _currentSheet);
					return true;
				case null:
					if (Form1.Report.ReportNull)
						Form1.Report.Add(pos.Col, pos.Row, Err.CellNull, "NULL", 0, _currentSheet);
					break;
				default:
					if (Form1.Report.ReportXData)
						Form1.Report.Add(
							pos.Col, pos.Row, Err.UnexpectedDataTypeError , "Check Spacings", 
							cell.GetType().ToString(), _currentSheet);
					break;
			}
			return false;
		}


        private bool CheckForLetter(object cell, Form1.Position pos) 
        {	switch (cell) {
				case string stringValue:
					if (stringValue.Any(char.IsLetter))
					{
						Form1.Report.Add(
							pos.Col, pos.Row, Err.IfLettersError, "String", stringValue, _currentSheet);
						return true;
					}
					else
					{
						return true;
					}

				case null:
					if (Form1.Report.ReportNull)
						Form1.Report.Add(pos.Col, pos.Row, Err.CellNull, "NULL", 0, _currentSheet);
					break;
				default:
					if (Form1.Report.ReportXData)
						Form1.Report.Add(
							pos.Col, pos.Row, Err.UnexpectedDataTypeError, "Check For Letter", 
							cell.GetType().ToString(), _currentSheet);
					break;
			}
			return false;
        }

        private bool CheckForNonAlpha(object cell, Form1.Position pos)
        {	switch (cell) {
				case string stringValue:
					foreach (char character in stringValue) {
						if (!char.IsLetterOrDigit(character) || char.IsWhiteSpace(character))
						{
							Form1.Report.Add(
								pos.Col, pos.Row, Err.NonAlpaError, "A-Z / 0-9", 
								stringValue, _currentSheet);
						}
					} return true;

				case null:
					if(Form1.Report.ReportNull)
						Form1.Report.Add(pos.Col, pos.Row, Err.CellNull, "A-Z / 0-9", 0, _currentSheet);
					break;
				default:
					if (Form1.Report.ReportXData)
						Form1.Report.Add(pos.Col, pos.Row, Err.UnexpectedDataTypeError, "Check for Non-Alpha",
							cell.GetType().ToString(), _currentSheet);
					break;
			}
			return false;
        }

        public bool CheckIfStringEqual (object cell, Form1.Position pos, string toCompare)
        {	switch(cell) {
				case string stringValue:
				if (toCompare.Equals(stringValue))
					return true;
				else {
					Form1.Report.Add(
						pos.Col, pos.Row, Err.StringEqual, toCompare, stringValue, _currentSheet);
					return false;
				}
				case null:
				if (Form1.Report.ReportNull)
					Form1.Report.Add(
						pos.Col, pos.Row, Err.CellNull, "Check String Equal", 0, _currentSheet);
				break;
				default:
				if (Form1.Report.ReportXData)
					Form1.Report.Add(
						pos.Col, pos.Row, Err.UnexpectedDataTypeError, "Check String Equal", 
						cell.GetType().ToString(), _currentSheet);
				break;
			}
            return false;
        }

        private bool CheckIfListEqual (string cell, Form1.Position pos, List<string> allowedWords)
        {	switch (cell) {

				case string stringValue:
					//TODO Look into. Doesn't make sense i think this should check for a match not (not match)
					int errors = 0;
					foreach (string str in allowedWords) 
					{
						if (!stringValue.Equals(str)) errors++;
					}
					if (errors >= allowedWords.Count) Form1.Report.Add(
						pos.Col, pos.Row, Err.StringEqual, "Array", stringValue, _currentSheet);
					break;

				case null:
					if (Form1.Report.ReportNull)
						Form1.Report.Add(pos.Col, pos.Row, Err.CellNull, "Check List", 0, _currentSheet);
					break;
				// default:
				// 	if (Form1.Report.ReportXData)
				// 		Form1.Report.Add(pos.Col, pos.Row, Err.UnexpectedDataTypeError, "Check List", 
				// 			cell.GetType().ToString(), _currentSheet);
				// 	break;
			}
			
            return true;
        }

        private bool CheckBeginsWith(object cell, Form1.Position pos, string startString)
        {	switch (cell) {
				case string stringValue:
					if (stringValue.StartsWith(startString))
						return false;
					else { Form1.Report.Add(
						pos.Col, pos.Row, Err.MustBeginWithError, startString, 
						stringValue, _currentSheet); return true; }

				case null:
					if (Form1.Report.ReportNull)
						Form1.Report.Add(
							pos.Col, pos.Row, Err.CellNull, "Check Begins With", 0, _currentSheet);
					break;
				default:
					if (Form1.Report.ReportXData)
						Form1.Report.Add(
							pos.Col, pos.Row, Err.UnexpectedDataTypeError, "Check Begins With", 
							cell.GetType().ToString(), _currentSheet);
					break;
			}
	        return false;
        }

        private bool CheckEndsWith (object cell, Form1.Position pos, string endString)
        {	switch (cell) {
				case string stringValue:
					if (stringValue.EndsWith(endString))
						return false;
					else
					{
						Form1.Report.Add(
							pos.Col, pos.Row, Err.MustEndWithError, endString, 
							stringValue, _currentSheet); return true;
					}

				case null:
					if (Form1.Report.ReportNull)
						Form1.Report.Add(
							pos.Col, pos.Row, Err.CellNull, "Check Ends With", 0, _currentSheet);
					break;
				default:
					if (Form1.Report.ReportXData)
						Form1.Report.Add(pos.Col, pos.Row, Err.UnexpectedDataTypeError, "Check Ends With", 
							cell.GetType().ToString(), _currentSheet);
					break;
			}
			return false;
		}

        private bool CheckLength(object cell, Form1.Position pos, int length)
        {	switch(cell){
				case string stringValue:
					if (stringValue.Length == length) { return true; }
					else
					{
						Form1.Report.Add(
							pos.Col, pos.Row, Err.MustHaveLenghtError, length, 
							stringValue.Length, _currentSheet); 
						return false;
					}

				case double doubleValue:
					if(doubleValue.ToString().Length == length) { return true; }
					else
					{
						Form1.Report.Add(
							pos.Col, pos.Row, Err.MustHaveLenghtError, length, 
							doubleValue.ToString().Length, _currentSheet); 
						return false;
					}
				case null:
					if (Form1.Report.ReportNull)
						Form1.Report.Add(
							pos.Col, pos.Row, Err.CellNull, "Check Length", 0, _currentSheet);
					break;
				default:
					if (Form1.Report.ReportXData)
						Form1.Report.Add(
							pos.Col, pos.Row, Err.UnexpectedDataTypeError, "Check Length", 
							cell.GetType().ToString(), _currentSheet);
					break;
			}
			return false;
        }

        private bool CheckNumber(object cell, Form1.Position pos, double checkAmount, bool checkType) // check type 0-lower 1 higher
		{	switch (cell) {
				case double doubleValue:
					if (checkType) {//check if higher
						if (doubleValue > checkAmount)
							Form1.Report.Add(
								pos.Col, pos.Row, Err.IfMoreThanError, $"> {checkAmount}", 
								doubleValue, _currentSheet);
					}
					else {
						if (doubleValue < checkAmount)
							Form1.Report.Add(
								pos.Col, pos.Row, Err.IfMoreThanError, $"< {checkAmount}", 
								doubleValue, _currentSheet);
					}
					break;
				case null:
					if (Form1.Report.ReportNull)
						Form1.Report.Add(
							pos.Col, pos.Row, Err.CellNull, "Check More than", 0, _currentSheet);
					break;
				default:
					if (Form1.Report.ReportXData)
						Form1.Report.Add(
							pos.Col, pos.Row, Err.UnexpectedDataTypeError, "Check More than", 
							cell.GetType().ToString(), _currentSheet);
					break;
			}
			return false;
		}
     
        public static NumberFormat GetNumberFormat(string cellFormat)
        {   const string generalNumberFormat = "0.00";
            const string currencyFormat = "$#,##0.00";
            const string percentageFormat = "0.00%";

            if (cellFormat.Equals(generalNumberFormat))
                return NumberFormat.Number;
            if (cellFormat.Equals(currencyFormat))
	            return NumberFormat.Currency;
            if (cellFormat.Equals(percentageFormat))
	            return NumberFormat.Percentage;
            return NumberFormat.Accounting;
        }

        public enum NumberFormat
        {   GeneralText,
            Accounting,
            Currency,
            Percentage,
            Number
        }
    }
}