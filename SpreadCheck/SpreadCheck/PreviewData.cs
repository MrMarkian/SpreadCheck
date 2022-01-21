using System;
using System.Globalization;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SpreadCheck
{
	public partial class PreviewData : Form
	{
		private readonly Excel.Worksheet _xlPreviewSheet;

		private readonly int _columnNumber;
	

		public PreviewData(Excel._Workbook xlWorkbook, int endPosition, int columnNo)
		{
			_xlPreviewSheet = (Excel.Worksheet)xlWorkbook.Worksheets.Item[Index: 1];
			_columnNumber = columnNo +1;
			
			InitializeComponent();
			MaxRowNumber.Maximum = endPosition+1;
		}

		private void RefreshBButton_Click(object sender, EventArgs e)
		{
			//double calling the interop is probably v.slow.. 
			PreviewList.Items.Clear();
			for (int count =1; count < MaxRowNumber.Value +1; count++) {
				
				switch (_xlPreviewSheet.Cells [count, _columnNumber].Value) {
					case string stringValue:
					PreviewList.Items.Add(stringValue);
					break;

					case double doubleValue:
					PreviewList.Items.Add(Convert.ToString(doubleValue, CultureInfo.InvariantCulture));
					break;

					case DateTime dateTimeValue:
					PreviewList.Items.Add(dateTimeValue.ToString(CultureInfo.InvariantCulture));
					break;

					case bool boolValue:
					PreviewList.Items.Add(boolValue.ToString());
					break;

					case null:
					PreviewList.Items.Add("Null");
					break;

					default:
					PreviewList.Items.Add("Unhandled Data Type");
					break;


				}



			}
		}

		private void PreviewList_VisibleChanged(object sender, EventArgs e)
		{
			RefreshBButton_Click(sender, e);
		}

		private void PreviewList_Resize(object sender, EventArgs e)
		{


		}
	}


}
