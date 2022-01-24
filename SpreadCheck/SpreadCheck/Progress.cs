using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SpreadCheck
{	public partial class Progress : Form
	{	public Progress()
		{	InitializeComponent();
			SetUpProgress();
		}
		private void StopButton_Click(object sender, EventArgs e)
		{	Form1.RulesRunning = false;
		}

		public void SetUpProgress()
		{
			ProgressLabel.Visible = false;
			StopButton.Visible = true;
		}

		public void ShowComplete()
		{
			RunProgress.Value = RunProgress.Maximum;
			StopButton.Visible = false;
			ProgressLabel.Visible = true;
		}
		
		public void UpdateProgressInfo(int row, int column, Stopwatch stopwatch, XLFunctions xlFunc, int errorNumbers)
		{
			RowCheckLabel.Text = $@"Row:{row}";
			ColumnCheckLabel.Text = $@"Column:{column}";
			UpdateProgressBar(row, column);
			ErrorsFoundLabel.Text = $@"Errors Found:{errorNumbers}";
			var time = TimeSpan.FromMilliseconds(stopwatch.ElapsedMilliseconds);
			string answer = $"{time.Hours:D2}h:{time.Minutes:D2}m:{time.Seconds:D2}s";
			ElapsedLabel.Text = $@"Time Elapsed:{answer}";
			FunctionsRunLabel.Text = $@"Functions Run:{xlFunc.FunctionCallCount.ToString()}";
			Application.DoEvents();
		}

		private void UpdateProgressBar(int row, int column)
		{
			if (RunProgress.Maximum == row)
			{
				RunProgress.Value = RunProgress.Maximum - 1;
			}
			else
			{
				RunProgress.Value = row;
			}
		}
		
		public void SetUpProgressBar(int enabledRuleListCount, int foundLastRow)
		{
			Show();
			RunProgress.Visible = true;
			RunProgress.Minimum = 0;
			// Set progress by number of rows x columns with enabled rule lists
			RunProgress.Maximum = foundLastRow;
		}
	}
}
