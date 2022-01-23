using System.ComponentModel;

namespace SpreadCheck
{
    partial class PickSheet
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.SheetsListBox = new System.Windows.Forms.ListBox();
            this.SelectButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // SheetsListBox
            // 
            this.SheetsListBox.FormattingEnabled = true;
            this.SheetsListBox.ItemHeight = 20;
            this.SheetsListBox.Location = new System.Drawing.Point(27, 31);
            this.SheetsListBox.Name = "SheetsListBox";
            this.SheetsListBox.Size = new System.Drawing.Size(430, 284);
            this.SheetsListBox.TabIndex = 0;
            this.SheetsListBox.SelectedIndexChanged += new System.EventHandler(this.SheetsListBox_SelectedIndexChanged);
            // 
            // SelectButton
            // 
            this.SelectButton.Location = new System.Drawing.Point(528, 223);
            this.SelectButton.Name = "SelectButton";
            this.SelectButton.Size = new System.Drawing.Size(218, 68);
            this.SelectButton.TabIndex = 1;
            this.SelectButton.Text = "Select";
            this.SelectButton.UseVisualStyleBackColor = true;
            this.SelectButton.Click += new System.EventHandler(this.SelectButton_Click);
            // 
            // PickSheet
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 358);
            this.Controls.Add(this.SelectButton);
            this.Controls.Add(this.SheetsListBox);
            this.Name = "PickSheet";
            this.Text = "PickSheet";
            this.Load += new System.EventHandler(this.PickSheet_Load);
            this.ResumeLayout(false);
        }

        private System.Windows.Forms.ListBox SheetsListBox;
        private System.Windows.Forms.Button SelectButton;

        #endregion
    }
}