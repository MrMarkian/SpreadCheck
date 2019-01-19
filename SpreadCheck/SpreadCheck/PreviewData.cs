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
	public partial class PreviewData : Form
	{
		public Excel.Workbook xlWrkBk;
		Excel.Worksheet xlPreviewSheet;

		public int ColumNumber;
	

		public PreviewData(Excel.Workbook xlWorkbook, Excel.Worksheet curSheet, int endPosition, int ColumnNo)
		{
			xlWrkBk = xlWorkbook;
			xlPreviewSheet = (Excel.Worksheet)xlWrkBk.Worksheets.get_Item(1);
			ColumNumber = ColumnNo +1;
			
			InitializeComponent();
			MaxRowNumber.Maximum = endPosition+1;
		}

		private void RefreshBButton_Click(object sender, EventArgs e)
		{
			//double calling the interop is probably v.slow.. 
			PreviewList.Items.Clear();
			for (int count =1; count < MaxRowNumber.Value +1; count++) {
				
				switch (xlPreviewSheet.Cells [count, ColumNumber].Value) {
					case string s:
					PreviewList.Items.Add(xlPreviewSheet.Cells [count, ColumNumber].Value);
					break;

					case double d:
					PreviewList.Items.Add(Convert.ToString(xlPreviewSheet.Cells [count, ColumNumber].Value));
					break;

					case DateTime dt:
					PreviewList.Items.Add(dt.ToString());
					break;

					case bool bl:
					PreviewList.Items.Add(bl.ToString());
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
