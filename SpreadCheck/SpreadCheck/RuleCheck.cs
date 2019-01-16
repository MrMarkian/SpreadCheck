using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace SpreadCheck
{
    class RuleCheck
    {
        private Excel.Worksheet currentSheet;
        public RuleCheck(Excel.Worksheet sheet) { currentSheet = sheet; }
      
		public static Reporter.ErrorMessages err = new Reporter.ErrorMessages();
	
        public bool CheckifEmpty(object Cell, Form1.Position pos) // complete
        {	switch (Cell) {
				case string s:
					if (String.IsNullOrWhiteSpace(s))
						Form1.Report.Add(pos.Col, pos.Row, err.EmptyError, "AnyThing", "WhiteSpace", currentSheet);
				break;
				case double d:
					if(d==0.00)
					Form1.Report.Add(pos.Col, pos.Row, err.EmptyError + " -Double", 1, 0, currentSheet);
				break;
				case null:
				if (Form1.Report.ReportNull)
					Form1.Report.Add(pos.Col, pos.Row, err.CellNull, 1, 0, currentSheet);
				break;
				default:
					if(Form1.Report.ReportXData)
					Form1.Report.Add(pos.Col, pos.Row, err.UnexpectedDataTypeError, "String", Cell.GetType().ToString(), currentSheet);
				break;
			}
			return false;
        }

        public bool CheckForNumbers (object Cell, Form1.Position pos) // complete
        {	switch (Cell) {
				case string s:
					if (s.Any(char.IsNumber)) 
					Form1.Report.Add(pos.Col, pos.Row, err.IfNumbersError + " -String", "String", (object)Cell, currentSheet);
					return true;
				case null:
				if (Form1.Report.ReportNull)
					Form1.Report.Add(pos.Col, pos.Row, err.CellNull, "A-Z", "NULL", currentSheet);
				break;

				default:
				if (Form1.Report.ReportXData)
					Form1.Report.Add(pos.Col, pos.Row, err.UnexpectedDataTypeError, "A-Z", Cell.GetType().ToString(), currentSheet);
				break;
			}
			return false;
        }

		public bool CheckForSpaces (object Cell, Form1.Position pos)
		{	switch (Cell) {
				case string s:
				if (s.Any(Char.IsWhiteSpace))
					Form1.Report.Add(pos.Col, pos.Row, err.SpacesError, "No Spaces", (object)Cell, currentSheet);
				return true;
				case null:
				if (Form1.Report.ReportNull)
					Form1.Report.Add(pos.Col, pos.Row, err.CellNull, "NULL", 0, currentSheet);
				break;
				default:
				if (Form1.Report.ReportXData)
					Form1.Report.Add(pos.Col, pos.Row, err.UnexpectedDataTypeError , "Check Spacings", Cell.GetType().ToString(), currentSheet);
				break;
			}
			return false;
		}


        public bool CheckForLetter(object Cell, Form1.Position pos) 
        {	switch (Cell) {
				case string s:
				if (s.Any(char.IsLetter))
					Form1.Report.Add(pos.Col, pos.Row, err.IfLettersError, "String", (object)Cell, currentSheet);
					return true;
				case null:
				if (Form1.Report.ReportNull)
					Form1.Report.Add(pos.Col, pos.Row, err.CellNull, "NULL", 0, currentSheet);
				break;
				default:
				if (Form1.Report.ReportXData)
					Form1.Report.Add(pos.Col, pos.Row, err.UnexpectedDataTypeError, "Check For Letter", Cell.GetType().ToString(), currentSheet);
				break;
			}
			return false;
        }

        public bool CheckForNonAlpha(object Cell, Form1.Position pos)
        {	switch (Cell) {
				case string s:

				foreach (char c in s) {
					if (!char.IsLetterOrDigit(c) || char.IsWhiteSpace(c)) { Form1.Report.Add(pos.Col, pos.Row, err.NonAlpaError, "A-Z / 0-9", s, currentSheet); }
				} return true;

				case null:
					if(Form1.Report.ReportNull)
					Form1.Report.Add(pos.Col, pos.Row, err.CellNull, "A-Z / 0-9", 0, currentSheet);
				break;
				default:
				if (Form1.Report.ReportXData)
					Form1.Report.Add(pos.Col, pos.Row, err.UnexpectedDataTypeError, "Check for Non-Alpha", Cell.GetType().ToString(), currentSheet);
				break;
			}
			return false;
        }

        public bool CheckifStringEqual (object Cell, Form1.Position pos, string toCompare)
        {	switch(Cell) {
				case string s:
				if (toCompare.Equals(s))
					return true;
				else {
					Form1.Report.Add(pos.Col, pos.Row, err.StringEqual, (object)toCompare, (object)Cell, currentSheet);
					return false;
				}
				case null:
				if (Form1.Report.ReportNull)
					Form1.Report.Add(pos.Col, pos.Row, err.CellNull, "Check String Equal", 0, currentSheet);
				break;
				default:
				if (Form1.Report.ReportXData)
					Form1.Report.Add(pos.Col, pos.Row, err.UnexpectedDataTypeError, "Check String Equal", Cell.GetType().ToString(), currentSheet);
				break;
			}
            return false;
        }

        public bool CheckifListEqual (string Cell, Form1.Position pos, List<string> allowedWords)
        {	switch (Cell) {

				case string s:
				int errors = 0;
				foreach (string str in allowedWords) {
					if (!s.Equals(str))
						errors++;
				}
				if (errors >= allowedWords.Count) Form1.Report.Add(pos.Col, pos.Row, err.StringEqual, "Array", (object)s, currentSheet);
				break;

				case null:
				if (Form1.Report.ReportNull)
					Form1.Report.Add(pos.Col, pos.Row, err.CellNull, "Check List", 0, currentSheet);
				break;
				default:
				if (Form1.Report.ReportXData)
					Form1.Report.Add(pos.Col, pos.Row, err.UnexpectedDataTypeError, "Check List", Cell.GetType().ToString(), currentSheet);
				break;
			}
			
            return true;
        }

        public bool CheckBeginsWith(object Cell, Form1.Position pos, string StartString)
        {	switch (Cell) {
				case string s:
				if (s.StartsWith(StartString))
					return false;
				else { Form1.Report.Add(pos.Col, pos.Row, err.MustBeginWithError, StartString, s, currentSheet); return true; }

				case null:
				if (Form1.Report.ReportNull)
					Form1.Report.Add(pos.Col, pos.Row, err.CellNull, "Check Begins With", 0, currentSheet);
				break;
				default:
				if (Form1.Report.ReportXData)
					Form1.Report.Add(pos.Col, pos.Row, err.UnexpectedDataTypeError, "Check Begins With", Cell.GetType().ToString(), currentSheet);
				break;
			}
				return false;
        }

        public bool CheckEndsWith (object Cell, Form1.Position pos, string EndString)
        {	switch (Cell) {
				case string s:
				if (s.EndsWith(EndString))
					return false;
				else { Form1.Report.Add(pos.Col, pos.Row, err.MustEndWithError, EndString, s, currentSheet); return true; }

				case null:
				if (Form1.Report.ReportNull)
					Form1.Report.Add(pos.Col, pos.Row, err.CellNull, "Check Ends With", 0, currentSheet);
				break;
				default:
				if (Form1.Report.ReportXData)
					Form1.Report.Add(pos.Col, pos.Row, err.UnexpectedDataTypeError, "Check Ends With", Cell.GetType().ToString(), currentSheet);
				break;
			}
			return false;
		}

        public bool CheckLength(object Cell, Form1.Position pos, int Length)
        {	switch(Cell){
				case string s:
				if (s.Length == Length) { return true; } else { Form1.Report.Add(pos.Col, pos.Row, err.MustHaveLenghtError, Length, s.Length, currentSheet); return false; }

				case double d:
					if(d.ToString().Length == Length) { return true; } else { Form1.Report.Add(pos.Col, pos.Row, err.MustHaveLenghtError, Length, d.ToString().Length, currentSheet); return false; }
				case null:
				if (Form1.Report.ReportNull)
					Form1.Report.Add(pos.Col, pos.Row, err.CellNull, "Check Length", 0, currentSheet);
				break;
				default:
				if (Form1.Report.ReportXData)
					Form1.Report.Add(pos.Col, pos.Row, err.UnexpectedDataTypeError, "Check Length", Cell.GetType().ToString(), currentSheet);
				break;
			}
			return false;
        }

		public bool CheckNumber(object Cell, Form1.Position pos, double checkamount, bool checktype) // check type 0-lower 1 higher
		{	switch (Cell) {
				case double d:
				if (checktype) {//check if higer
					if (d > checkamount)
						Form1.Report.Add(pos.Col, pos.Row, err.IfMoreThanError, "> " + checkamount.ToString(), d, currentSheet);
				}
				else {
					if (d < checkamount)
						Form1.Report.Add(pos.Col, pos.Row, err.IfMoreThanError, "< " + checkamount.ToString(), d, currentSheet);
				}
				break;
				case null:
					if (Form1.Report.ReportNull)
					Form1.Report.Add(pos.Col, pos.Row, err.CellNull, "Check More than", 0, currentSheet);
				break;
				default:
				if (Form1.Report.ReportXData)
					Form1.Report.Add(pos.Col, pos.Row, err.UnexpectedDataTypeError, "Check More than", Cell.GetType().ToString(), currentSheet);
				break;
			}
			return false;
		}
     
        public static NumberFormat GetNumberFormat(string cellFormat)
        {   string GeneralNumberFormat = "0.00";
            string CurrencyFormat = "$#,##0.00";
            string PercentageFormat = "0.00%";

            if (cellFormat.Equals(GeneralNumberFormat))
                return NumberFormat.Number;
            else if (cellFormat.Equals(CurrencyFormat))
                return NumberFormat.Currency;
            else if (cellFormat.Equals(PercentageFormat))
                return NumberFormat.Percentage;
            else
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