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
		{	CellEmptyErrorText.Text = RuleCheck.Err.EmptyError;
			CellContainsSpacesError.Text = RuleCheck.Err.SpacesError;
			CellContainsNonAlphaError.Text = RuleCheck.Err.NonAlpaError;
			CellContainsNumbersError.Text = RuleCheck.Err.IfNumbersError;
			CellContainsLettersError.Text = RuleCheck.Err.IfLettersError;
			CellMustBeginWithError.Text = RuleCheck.Err.MustBeginWithError;
			CellMustEndWithError.Text = RuleCheck.Err.MustEndWithError;
			CellLengthError.Text = RuleCheck.Err.MustHaveLenghtError;
			CellAllowedItemsError.Text = RuleCheck.Err.AllowedItemsListError;
			CellMoreThanError.Text = RuleCheck.Err.IfMoreThanError;
			CellLessThanError.Text = RuleCheck.Err.IfLessThanError;
			CellStringEqualError.Text = RuleCheck.Err.StringEqual;
			CellNullError.Text = RuleCheck.Err.CellNull;
			CellUnexpectedError.Text = RuleCheck.Err.UnexpectedDataTypeError;
			HyperLinkEnabledCheckbox.Checked = Form1.Report.IncludeHyperLink;
			IncludeOriginalData.Checked = Form1.Report.IncludeOriginalData;
			IncludeEnabledColumns.Checked = Form1.Report.IncludeEnabledData;
			IncludeFullRow.Checked = Form1.Report.IncludeFullRowOfData;
			IncludePivotCheckBox.Checked = Form1.Report.IncludePivotTable;
			EnumLabel.Text = "Cell BackColor on Error: " + RuleCheck.Err.BackCellColor.ToString();
			NullReportCheckbox.Checked = Form1.Report.ReportNull;
			InvalidTypesReportCheckBox.Checked = Form1.Report.ReportXData;
			UseColumnNumberRadio.Checked = Form1.Report.UseHeaderNames;
			foreach (string ColStrt in Enum.GetNames(typeof(Excel.XlRgbColor)))
			{	ErrorColorCombo.Items.Add(ColStrt);		}
			ErrorColorCombo.SelectedIndex = ErrorColorCombo.FindStringExact(RuleCheck.Err.BackCellColor.ToString());
		}

		private void ErrorColorCombo_SelectedIndexChanged(object sender, EventArgs e)
		{
			RuleCheck.Err.BackCellColor = (Excel.XlRgbColor)Enum.Parse(typeof(Excel.XlRgbColor), ErrorColorCombo.SelectedItem.ToString());
			EnumLabel.Text = "Cell BackColor on Error: " + RuleCheck.Err.BackCellColor.ToString();
			
		}

		private void radioButton2_CheckedChanged(object sender, EventArgs e)
		{		}

		private void SaveButton_Click(object sender, EventArgs e)
		{
			RuleCheck.Err.EmptyError = CellEmptyErrorText.Text;
			RuleCheck.Err.SpacesError = CellContainsSpacesError.Text;
			RuleCheck.Err.NonAlpaError = CellContainsNonAlphaError.Text;
			RuleCheck.Err.IfNumbersError = CellContainsNumbersError.Text;
			RuleCheck.Err.IfLettersError = CellContainsLettersError.Text;
			RuleCheck.Err.MustBeginWithError = CellMustBeginWithError.Text;
			RuleCheck.Err.MustEndWithError = CellMustEndWithError.Text;
			RuleCheck.Err.MustHaveLenghtError = CellLengthError.Text;
			RuleCheck.Err.AllowedItemsListError = CellAllowedItemsError.Text;
			RuleCheck.Err.IfMoreThanError = CellMoreThanError.Text;
			RuleCheck.Err.IfLessThanError = CellLessThanError.Text;
			RuleCheck.Err.StringEqual = CellStringEqualError.Text;
			RuleCheck.Err.CellNull = CellNullError.Text;
			RuleCheck.Err.UnexpectedDataTypeError = CellUnexpectedError.Text;
			Form1.Report.IncludeHyperLink = HyperLinkEnabledCheckbox.Checked;
			Form1.Report.IncludeOriginalData = IncludeOriginalData.Checked;
			Form1.Report.IncludeFullRowOfData = IncludeFullRow.Checked;
			Form1.Report.IncludeEnabledData = IncludeEnabledColumns.Checked;
			Form1.Report.UseHeaderNames = UseColumnNumberRadio.Checked;
			Form1.Report.IncludePivotTable = IncludePivotCheckBox.Checked;
			Form1.Report.ReportNull = NullReportCheckbox.Checked;
			Form1.Report.ReportXData = InvalidTypesReportCheckBox.Checked;
			RuleCheck.Err.BackCellColor = (Excel.XlRgbColor)Enum.Parse(typeof(Excel.XlRgbColor), ErrorColorCombo.Items [ErrorColorCombo.SelectedIndex].ToString());
			ReportingOptions.ActiveForm.Close();
		}
	}
}
