using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SpreadCheck
{
	public partial class ReportingOptions : Form
	{
		public ReportingOptions()
		{	InitializeComponent();		}

		private void ReportingOptions_Load(object sender, EventArgs e)
		{	CellEmptyErrorText.Text = RuleCheck.err.EmptyError;
			CellContainsSpacesError.Text = RuleCheck.err.SpacesError;
			CellContainsNonAlphaError.Text = RuleCheck.err.NonAlpaError;
			CellContainsNumbersError.Text = RuleCheck.err.IfNumbersError;
			CellContainsLettersError.Text = RuleCheck.err.IfLettersError;
			CellMustBeginWithError.Text = RuleCheck.err.MustBeginWithError;
			CellMustEndWithError.Text = RuleCheck.err.MustEndWithError;
			CellLengthError.Text = RuleCheck.err.MustHaveLenghtError;
			CellAllowedItemsError.Text = RuleCheck.err.AllowedItemsListError;
			CellMoreThanError.Text = RuleCheck.err.IfMoreThanError;
			CellLessThanError.Text = RuleCheck.err.IfLessThanError;
			CellStringEqualError.Text = RuleCheck.err.StringEqual;
			CellNullError.Text = RuleCheck.err.CellNull;
			CellUnexpectedError.Text = RuleCheck.err.UnexpectedDataTypeError;
			HyperLinkEnabledCheckbox.Checked = Form1.Report.IncludeHyperLink;
			IncludeOriginalData.Checked = Form1.Report.IncludeOriginalData;
			IncludeEnabledColumns.Checked = Form1.Report.IncludeEnabledData;
			IncludeFullRow.Checked = Form1.Report.IncludeFullRowOfData;
			IncludePivotCheckBox.Checked = Form1.Report.IncludePivotTable;
			EnumLabel.Text = "Cell BackColor on Error: " + RuleCheck.err.BackCellColor.ToString();
			NullReportCheckbox.Checked = Form1.Report.ReportNull;
			InvalidTypesReportCheckBox.Checked = Form1.Report.ReportXData;
			UseColumnNumberRadio.Checked = Form1.Report.UseHeaderNames;
			foreach (string ColStrt in Enum.GetNames(typeof(Excel.XlRgbColor)))
			{	ErrorColorCombo.Items.Add(ColStrt);		}
			ErrorColorCombo.SelectedIndex = ErrorColorCombo.FindStringExact(RuleCheck.err.BackCellColor.ToString());
		}

		private void ErrorColorCombo_SelectedIndexChanged(object sender, EventArgs e)
		{
			RuleCheck.err.BackCellColor = (Excel.XlRgbColor)Enum.Parse(typeof(Excel.XlRgbColor), ErrorColorCombo.SelectedItem.ToString());
			EnumLabel.Text = "Cell BackColor on Error: " + RuleCheck.err.BackCellColor.ToString();
			
		}

		private void radioButton2_CheckedChanged(object sender, EventArgs e)
		{		}

		private void SaveButton_Click(object sender, EventArgs e)
		{
			RuleCheck.err.EmptyError = CellEmptyErrorText.Text;
			RuleCheck.err.SpacesError = CellContainsSpacesError.Text;
			RuleCheck.err.NonAlpaError = CellContainsNonAlphaError.Text;
			RuleCheck.err.IfNumbersError = CellContainsNumbersError.Text;
			RuleCheck.err.IfLettersError = CellContainsLettersError.Text;
			RuleCheck.err.MustBeginWithError = CellMustBeginWithError.Text;
			RuleCheck.err.MustEndWithError = CellMustEndWithError.Text;
			RuleCheck.err.MustHaveLenghtError = CellLengthError.Text;
			RuleCheck.err.AllowedItemsListError = CellAllowedItemsError.Text;
			RuleCheck.err.IfMoreThanError = CellMoreThanError.Text;
			RuleCheck.err.IfLessThanError = CellLessThanError.Text;
			RuleCheck.err.StringEqual = CellStringEqualError.Text;
			RuleCheck.err.CellNull = CellNullError.Text;
			RuleCheck.err.UnexpectedDataTypeError = CellUnexpectedError.Text;
			Form1.Report.IncludeHyperLink = HyperLinkEnabledCheckbox.Checked;
			Form1.Report.IncludeOriginalData = IncludeOriginalData.Checked;
			Form1.Report.IncludeFullRowOfData = IncludeFullRow.Checked;
			Form1.Report.IncludeEnabledData = IncludeEnabledColumns.Checked;
			Form1.Report.UseHeaderNames = UseColumnNumberRadio.Checked;
			Form1.Report.IncludePivotTable = IncludePivotCheckBox.Checked;
			Form1.Report.ReportNull = NullReportCheckbox.Checked;
			Form1.Report.ReportXData = InvalidTypesReportCheckBox.Checked;
			RuleCheck.err.BackCellColor = (Excel.XlRgbColor)Enum.Parse(typeof(Excel.XlRgbColor), ErrorColorCombo.Items [ErrorColorCombo.SelectedIndex].ToString());
			ReportingOptions.ActiveForm.Close();
		}
	}
}
