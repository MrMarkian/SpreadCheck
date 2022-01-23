using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace SpreadCheck
{
    public partial class PickSheet : Form
    {
        public int SheetIndexSelected { get; private set; }

        
        public PickSheet(int sheetIndexSelected, List<string> sheetNames)
        {
            SheetIndexSelected = sheetIndexSelected;
            InitializeComponent();
            for (int index = 0; index < sheetNames.Count; index++)
            {
                string sheet = sheetNames[index];
                SheetsListBox.Items.Add(sheet);
                
            }
        }

        private void SelectButton_Click(object sender, EventArgs e)
        {
            SheetIndexSelected = SheetsListBox.SelectedIndex;
            MessageBox.Show($@"selected sheet index = {SheetIndexSelected}");
            DialogResult = DialogResult.OK;
            Close();
        }
        private void PickSheet_Load(object sender, EventArgs e)
        {
            
        }

        private void SheetsListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sheet = SheetsListBox.Items[SheetsListBox.SelectedIndex].ToString();
            if (sheet == "Errors")
            {
                MessageBox.Show(@"Errors sheet cannot be selected try closing the program and renaming the sheet");
                SheetsListBox.SelectedIndex = 0;
            }
        }
    }
}