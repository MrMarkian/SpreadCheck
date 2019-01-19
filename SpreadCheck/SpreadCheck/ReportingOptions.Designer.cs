namespace SpreadCheck
{
	partial class ReportingOptions
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

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
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.IncludeFullRow = new System.Windows.Forms.RadioButton();
			this.IncludeEnabledColumns = new System.Windows.Forms.RadioButton();
			this.IncludeOriginalData = new System.Windows.Forms.CheckBox();
			this.CellEmptyErrorText = new System.Windows.Forms.TextBox();
			this.CellContainsSpacesError = new System.Windows.Forms.TextBox();
			this.CellContainsNonAlphaError = new System.Windows.Forms.TextBox();
			this.CellContainsNumbersError = new System.Windows.Forms.TextBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label14 = new System.Windows.Forms.Label();
			this.CellUnexpectedError = new System.Windows.Forms.TextBox();
			this.label13 = new System.Windows.Forms.Label();
			this.CellNullError = new System.Windows.Forms.TextBox();
			this.label12 = new System.Windows.Forms.Label();
			this.CellStringEqualError = new System.Windows.Forms.TextBox();
			this.label11 = new System.Windows.Forms.Label();
			this.CellLessThanError = new System.Windows.Forms.TextBox();
			this.label10 = new System.Windows.Forms.Label();
			this.CellMoreThanError = new System.Windows.Forms.TextBox();
			this.label9 = new System.Windows.Forms.Label();
			this.CellAllowedItemsError = new System.Windows.Forms.TextBox();
			this.label8 = new System.Windows.Forms.Label();
			this.CellLengthError = new System.Windows.Forms.TextBox();
			this.label7 = new System.Windows.Forms.Label();
			this.CellMustEndWithError = new System.Windows.Forms.TextBox();
			this.label6 = new System.Windows.Forms.Label();
			this.CellMustBeginWithError = new System.Windows.Forms.TextBox();
			this.label5 = new System.Windows.Forms.Label();
			this.CellContainsLettersError = new System.Windows.Forms.TextBox();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.HyperLinkEnabledCheckbox = new System.Windows.Forms.CheckBox();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.UseColumnNumberRadio = new System.Windows.Forms.RadioButton();
			this.UseHeaderNamesRadio = new System.Windows.Forms.RadioButton();
			this.colorDialog = new System.Windows.Forms.ColorDialog();
			this.ErrorColorCombo = new System.Windows.Forms.ComboBox();
			this.EnumLabel = new System.Windows.Forms.Label();
			this.SaveButton = new System.Windows.Forms.Button();
			this.IncludePivotCheckBox = new System.Windows.Forms.CheckBox();
			this.NullReportCheckbox = new System.Windows.Forms.CheckBox();
			this.InvalidTypesReportCheckBox = new System.Windows.Forms.CheckBox();
			this.statusStrip1 = new System.Windows.Forms.StatusStrip();
			this.RuleSetLocation = new System.Windows.Forms.ToolStripStatusLabel();
			this.groupBox1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.groupBox3.SuspendLayout();
			this.statusStrip1.SuspendLayout();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.IncludeFullRow);
			this.groupBox1.Controls.Add(this.IncludePivotCheckBox);
			this.groupBox1.Controls.Add(this.InvalidTypesReportCheckBox);
			this.groupBox1.Controls.Add(this.IncludeEnabledColumns);
			this.groupBox1.Controls.Add(this.NullReportCheckbox);
			this.groupBox1.Controls.Add(this.IncludeOriginalData);
			this.groupBox1.Controls.Add(this.HyperLinkEnabledCheckbox);
			this.groupBox1.Location = new System.Drawing.Point(12, 12);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(227, 194);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Included Data / Columns";
			// 
			// IncludeFullRow
			// 
			this.IncludeFullRow.AutoSize = true;
			this.IncludeFullRow.Location = new System.Drawing.Point(19, 66);
			this.IncludeFullRow.Name = "IncludeFullRow";
			this.IncludeFullRow.Size = new System.Drawing.Size(134, 17);
			this.IncludeFullRow.TabIndex = 2;
			this.IncludeFullRow.Text = "Include full row of Data";
			this.IncludeFullRow.UseVisualStyleBackColor = true;
			// 
			// IncludeEnabledColumns
			// 
			this.IncludeEnabledColumns.AutoSize = true;
			this.IncludeEnabledColumns.Checked = true;
			this.IncludeEnabledColumns.Location = new System.Drawing.Point(19, 43);
			this.IncludeEnabledColumns.Name = "IncludeEnabledColumns";
			this.IncludeEnabledColumns.Size = new System.Drawing.Size(169, 17);
			this.IncludeEnabledColumns.TabIndex = 1;
			this.IncludeEnabledColumns.TabStop = true;
			this.IncludeEnabledColumns.Text = "Only Include Enabled Columns";
			this.IncludeEnabledColumns.UseVisualStyleBackColor = true;
			// 
			// IncludeOriginalData
			// 
			this.IncludeOriginalData.AutoSize = true;
			this.IncludeOriginalData.Checked = true;
			this.IncludeOriginalData.CheckState = System.Windows.Forms.CheckState.Checked;
			this.IncludeOriginalData.Location = new System.Drawing.Point(7, 20);
			this.IncludeOriginalData.Name = "IncludeOriginalData";
			this.IncludeOriginalData.Size = new System.Drawing.Size(209, 17);
			this.IncludeOriginalData.TabIndex = 0;
			this.IncludeOriginalData.Text = "Inlcude a copy of original data in report";
			this.IncludeOriginalData.UseVisualStyleBackColor = true;
			// 
			// CellEmptyErrorText
			// 
			this.CellEmptyErrorText.Location = new System.Drawing.Point(14, 40);
			this.CellEmptyErrorText.Name = "CellEmptyErrorText";
			this.CellEmptyErrorText.Size = new System.Drawing.Size(200, 20);
			this.CellEmptyErrorText.TabIndex = 1;
			// 
			// CellContainsSpacesError
			// 
			this.CellContainsSpacesError.Location = new System.Drawing.Point(14, 79);
			this.CellContainsSpacesError.Name = "CellContainsSpacesError";
			this.CellContainsSpacesError.Size = new System.Drawing.Size(200, 20);
			this.CellContainsSpacesError.TabIndex = 2;
			// 
			// CellContainsNonAlphaError
			// 
			this.CellContainsNonAlphaError.Location = new System.Drawing.Point(14, 132);
			this.CellContainsNonAlphaError.Name = "CellContainsNonAlphaError";
			this.CellContainsNonAlphaError.Size = new System.Drawing.Size(200, 20);
			this.CellContainsNonAlphaError.TabIndex = 3;
			// 
			// CellContainsNumbersError
			// 
			this.CellContainsNumbersError.Location = new System.Drawing.Point(11, 178);
			this.CellContainsNumbersError.Name = "CellContainsNumbersError";
			this.CellContainsNumbersError.Size = new System.Drawing.Size(200, 20);
			this.CellContainsNumbersError.TabIndex = 4;
			// 
			// groupBox2
			// 
			this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox2.AutoSize = true;
			this.groupBox2.Controls.Add(this.label14);
			this.groupBox2.Controls.Add(this.CellUnexpectedError);
			this.groupBox2.Controls.Add(this.label13);
			this.groupBox2.Controls.Add(this.CellNullError);
			this.groupBox2.Controls.Add(this.label12);
			this.groupBox2.Controls.Add(this.CellStringEqualError);
			this.groupBox2.Controls.Add(this.label11);
			this.groupBox2.Controls.Add(this.CellLessThanError);
			this.groupBox2.Controls.Add(this.label10);
			this.groupBox2.Controls.Add(this.CellMoreThanError);
			this.groupBox2.Controls.Add(this.label9);
			this.groupBox2.Controls.Add(this.CellAllowedItemsError);
			this.groupBox2.Controls.Add(this.label8);
			this.groupBox2.Controls.Add(this.CellLengthError);
			this.groupBox2.Controls.Add(this.label7);
			this.groupBox2.Controls.Add(this.CellMustEndWithError);
			this.groupBox2.Controls.Add(this.label6);
			this.groupBox2.Controls.Add(this.CellMustBeginWithError);
			this.groupBox2.Controls.Add(this.label5);
			this.groupBox2.Controls.Add(this.CellContainsLettersError);
			this.groupBox2.Controls.Add(this.label4);
			this.groupBox2.Controls.Add(this.label3);
			this.groupBox2.Controls.Add(this.label2);
			this.groupBox2.Controls.Add(this.label1);
			this.groupBox2.Controls.Add(this.CellEmptyErrorText);
			this.groupBox2.Controls.Add(this.CellContainsNumbersError);
			this.groupBox2.Controls.Add(this.CellContainsSpacesError);
			this.groupBox2.Controls.Add(this.CellContainsNonAlphaError);
			this.groupBox2.Location = new System.Drawing.Point(245, 12);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(428, 364);
			this.groupBox2.TabIndex = 5;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Error Strings";
			// 
			// label14
			// 
			this.label14.AutoSize = true;
			this.label14.Location = new System.Drawing.Point(219, 306);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(143, 13);
			this.label14.TabIndex = 28;
			this.label14.Text = "Unexpected DataType  Error";
			// 
			// CellUnexpectedError
			// 
			this.CellUnexpectedError.Location = new System.Drawing.Point(219, 322);
			this.CellUnexpectedError.Name = "CellUnexpectedError";
			this.CellUnexpectedError.Size = new System.Drawing.Size(200, 20);
			this.CellUnexpectedError.TabIndex = 27;
			// 
			// label13
			// 
			this.label13.AutoSize = true;
			this.label13.Location = new System.Drawing.Point(11, 306);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(70, 13);
			this.label13.TabIndex = 26;
			this.label13.Text = "Cell Null Error";
			// 
			// CellNullError
			// 
			this.CellNullError.Location = new System.Drawing.Point(11, 322);
			this.CellNullError.Name = "CellNullError";
			this.CellNullError.Size = new System.Drawing.Size(200, 20);
			this.CellNullError.TabIndex = 25;
			// 
			// label12
			// 
			this.label12.AutoSize = true;
			this.label12.Location = new System.Drawing.Point(216, 259);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(89, 13);
			this.label12.TabIndex = 24;
			this.label12.Text = "String Equal Error";
			// 
			// CellStringEqualError
			// 
			this.CellStringEqualError.Location = new System.Drawing.Point(219, 275);
			this.CellStringEqualError.Name = "CellStringEqualError";
			this.CellStringEqualError.Size = new System.Drawing.Size(200, 20);
			this.CellStringEqualError.TabIndex = 23;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Location = new System.Drawing.Point(219, 216);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(91, 13);
			this.label11.TabIndex = 22;
			this.label11.Text = "If Less Than Error";
			// 
			// CellLessThanError
			// 
			this.CellLessThanError.Location = new System.Drawing.Point(219, 232);
			this.CellLessThanError.Name = "CellLessThanError";
			this.CellLessThanError.Size = new System.Drawing.Size(200, 20);
			this.CellLessThanError.TabIndex = 21;
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Location = new System.Drawing.Point(219, 162);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(93, 13);
			this.label10.TabIndex = 20;
			this.label10.Text = "If More Than Error";
			// 
			// CellMoreThanError
			// 
			this.CellMoreThanError.Location = new System.Drawing.Point(219, 178);
			this.CellMoreThanError.Name = "CellMoreThanError";
			this.CellMoreThanError.Size = new System.Drawing.Size(200, 20);
			this.CellMoreThanError.TabIndex = 19;
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Location = new System.Drawing.Point(222, 116);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(97, 13);
			this.label9.TabIndex = 18;
			this.label9.Text = "Allowed Items Error";
			// 
			// CellAllowedItemsError
			// 
			this.CellAllowedItemsError.Location = new System.Drawing.Point(222, 132);
			this.CellAllowedItemsError.Name = "CellAllowedItemsError";
			this.CellAllowedItemsError.Size = new System.Drawing.Size(200, 20);
			this.CellAllowedItemsError.TabIndex = 17;
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(222, 63);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(132, 13);
			this.label8.TabIndex = 16;
			this.label8.Text = "Must Have Length of Error";
			// 
			// CellLengthError
			// 
			this.CellLengthError.Location = new System.Drawing.Point(222, 79);
			this.CellLengthError.Name = "CellLengthError";
			this.CellLengthError.Size = new System.Drawing.Size(200, 20);
			this.CellLengthError.TabIndex = 15;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(222, 24);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(102, 13);
			this.label7.TabIndex = 14;
			this.label7.Text = "Must End With Error";
			// 
			// CellMustEndWithError
			// 
			this.CellMustEndWithError.Location = new System.Drawing.Point(222, 40);
			this.CellMustEndWithError.Name = "CellMustEndWithError";
			this.CellMustEndWithError.Size = new System.Drawing.Size(200, 20);
			this.CellMustEndWithError.TabIndex = 13;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(11, 259);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(110, 13);
			this.label6.TabIndex = 12;
			this.label6.Text = "Must Begin With Error";
			// 
			// CellMustBeginWithError
			// 
			this.CellMustBeginWithError.Location = new System.Drawing.Point(11, 275);
			this.CellMustBeginWithError.Name = "CellMustBeginWithError";
			this.CellMustBeginWithError.Size = new System.Drawing.Size(200, 20);
			this.CellMustBeginWithError.TabIndex = 11;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(8, 216);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(124, 13);
			this.label5.TabIndex = 10;
			this.label5.Text = "Warn If Contains Letters:";
			// 
			// CellContainsLettersError
			// 
			this.CellContainsLettersError.Location = new System.Drawing.Point(11, 232);
			this.CellContainsLettersError.Name = "CellContainsLettersError";
			this.CellContainsLettersError.Size = new System.Drawing.Size(200, 20);
			this.CellContainsLettersError.TabIndex = 9;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(11, 162);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(134, 13);
			this.label4.TabIndex = 8;
			this.label4.Text = "Warn If Contains Numbers:";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(14, 116);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(181, 13);
			this.label3.TabIndex = 7;
			this.label3.Text = "Warn If Contains Non AlphaNumeric:";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(14, 63);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(128, 13);
			this.label2.TabIndex = 6;
			this.label2.Text = "Warn If Contains Spaces:";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(14, 24);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(77, 13);
			this.label1.TabIndex = 5;
			this.label1.Text = "Warn If Empty:";
			// 
			// HyperLinkEnabledCheckbox
			// 
			this.HyperLinkEnabledCheckbox.AutoSize = true;
			this.HyperLinkEnabledCheckbox.ForeColor = System.Drawing.Color.Red;
			this.HyperLinkEnabledCheckbox.Location = new System.Drawing.Point(7, 135);
			this.HyperLinkEnabledCheckbox.Name = "HyperLinkEnabledCheckbox";
			this.HyperLinkEnabledCheckbox.Size = new System.Drawing.Size(190, 17);
			this.HyperLinkEnabledCheckbox.TabIndex = 3;
			this.HyperLinkEnabledCheckbox.Text = "Inlcude a HyperLink to Errored Cell";
			this.HyperLinkEnabledCheckbox.UseVisualStyleBackColor = true;
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.UseColumnNumberRadio);
			this.groupBox3.Controls.Add(this.UseHeaderNamesRadio);
			this.groupBox3.Location = new System.Drawing.Point(12, 217);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(227, 67);
			this.groupBox3.TabIndex = 3;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Column Error Type";
			// 
			// UseColumnNumberRadio
			// 
			this.UseColumnNumberRadio.AutoSize = true;
			this.UseColumnNumberRadio.Location = new System.Drawing.Point(19, 44);
			this.UseColumnNumberRadio.Name = "UseColumnNumberRadio";
			this.UseColumnNumberRadio.Size = new System.Drawing.Size(122, 17);
			this.UseColumnNumberRadio.TabIndex = 2;
			this.UseColumnNumberRadio.Text = "Use Column Number";
			this.UseColumnNumberRadio.UseVisualStyleBackColor = true;
			// 
			// UseHeaderNamesRadio
			// 
			this.UseHeaderNamesRadio.AutoSize = true;
			this.UseHeaderNamesRadio.Checked = true;
			this.UseHeaderNamesRadio.Location = new System.Drawing.Point(19, 21);
			this.UseHeaderNamesRadio.Name = "UseHeaderNamesRadio";
			this.UseHeaderNamesRadio.Size = new System.Drawing.Size(110, 17);
			this.UseHeaderNamesRadio.TabIndex = 1;
			this.UseHeaderNamesRadio.TabStop = true;
			this.UseHeaderNamesRadio.Text = "Use HeaderName";
			this.UseHeaderNamesRadio.UseVisualStyleBackColor = true;
			this.UseHeaderNamesRadio.CheckedChanged += new System.EventHandler(this.radioButton2_CheckedChanged);
			// 
			// ErrorColorCombo
			// 
			this.ErrorColorCombo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.ErrorColorCombo.FormattingEnabled = true;
			this.ErrorColorCombo.Location = new System.Drawing.Point(12, 310);
			this.ErrorColorCombo.Name = "ErrorColorCombo";
			this.ErrorColorCombo.Size = new System.Drawing.Size(209, 21);
			this.ErrorColorCombo.TabIndex = 6;
			this.ErrorColorCombo.SelectedIndexChanged += new System.EventHandler(this.ErrorColorCombo_SelectedIndexChanged);
			// 
			// EnumLabel
			// 
			this.EnumLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.EnumLabel.AutoSize = true;
			this.EnumLabel.Location = new System.Drawing.Point(12, 294);
			this.EnumLabel.Name = "EnumLabel";
			this.EnumLabel.Size = new System.Drawing.Size(122, 13);
			this.EnumLabel.TabIndex = 7;
			this.EnumLabel.Text = "Cell BackColor on Error: ";
			// 
			// SaveButton
			// 
			this.SaveButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.SaveButton.Location = new System.Drawing.Point(12, 337);
			this.SaveButton.Name = "SaveButton";
			this.SaveButton.Size = new System.Drawing.Size(209, 33);
			this.SaveButton.TabIndex = 8;
			this.SaveButton.Text = "Save";
			this.SaveButton.UseVisualStyleBackColor = true;
			this.SaveButton.Click += new System.EventHandler(this.SaveButton_Click);
			// 
			// IncludePivotCheckBox
			// 
			this.IncludePivotCheckBox.AutoSize = true;
			this.IncludePivotCheckBox.ForeColor = System.Drawing.Color.Red;
			this.IncludePivotCheckBox.Location = new System.Drawing.Point(7, 158);
			this.IncludePivotCheckBox.Name = "IncludePivotCheckBox";
			this.IncludePivotCheckBox.Size = new System.Drawing.Size(170, 17);
			this.IncludePivotCheckBox.TabIndex = 9;
			this.IncludePivotCheckBox.Text = "Inlcude a PivotTable in Report";
			this.IncludePivotCheckBox.UseVisualStyleBackColor = true;
			// 
			// NullReportCheckbox
			// 
			this.NullReportCheckbox.AutoSize = true;
			this.NullReportCheckbox.Location = new System.Drawing.Point(7, 89);
			this.NullReportCheckbox.Name = "NullReportCheckbox";
			this.NullReportCheckbox.Size = new System.Drawing.Size(164, 17);
			this.NullReportCheckbox.TabIndex = 10;
			this.NullReportCheckbox.Text = "Report Null Values in Checks";
			this.NullReportCheckbox.UseVisualStyleBackColor = true;
			// 
			// InvalidTypesReportCheckBox
			// 
			this.InvalidTypesReportCheckBox.AutoSize = true;
			this.InvalidTypesReportCheckBox.Location = new System.Drawing.Point(7, 112);
			this.InvalidTypesReportCheckBox.Name = "InvalidTypesReportCheckBox";
			this.InvalidTypesReportCheckBox.Size = new System.Drawing.Size(147, 17);
			this.InvalidTypesReportCheckBox.TabIndex = 11;
			this.InvalidTypesReportCheckBox.Text = "Report Invalid DataTypes";
			this.InvalidTypesReportCheckBox.UseVisualStyleBackColor = true;
			// 
			// statusStrip1
			// 
			this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.RuleSetLocation});
			this.statusStrip1.Location = new System.Drawing.Point(0, 379);
			this.statusStrip1.Name = "statusStrip1";
			this.statusStrip1.Size = new System.Drawing.Size(674, 22);
			this.statusStrip1.SizingGrip = false;
			this.statusStrip1.TabIndex = 12;
			this.statusStrip1.Text = "statusStrip1";
			// 
			// RuleSetLocation
			// 
			this.RuleSetLocation.Name = "RuleSetLocation";
			this.RuleSetLocation.Size = new System.Drawing.Size(98, 17);
			this.RuleSetLocation.Text = "RuleSet Location:";
			// 
			// ReportingOptions
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(674, 401);
			this.Controls.Add(this.statusStrip1);
			this.Controls.Add(this.SaveButton);
			this.Controls.Add(this.EnumLabel);
			this.Controls.Add(this.ErrorColorCombo);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.MinimumSize = new System.Drawing.Size(690, 440);
			this.Name = "ReportingOptions";
			this.ShowIcon = false;
			this.ShowInTaskbar = false;
			this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Reporting Options";
			this.TopMost = true;
			this.Load += new System.EventHandler(this.ReportingOptions_Load);
			this.groupBox1.ResumeLayout(false);
			this.groupBox1.PerformLayout();
			this.groupBox2.ResumeLayout(false);
			this.groupBox2.PerformLayout();
			this.groupBox3.ResumeLayout(false);
			this.groupBox3.PerformLayout();
			this.statusStrip1.ResumeLayout(false);
			this.statusStrip1.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.RadioButton IncludeFullRow;
		private System.Windows.Forms.RadioButton IncludeEnabledColumns;
		private System.Windows.Forms.CheckBox IncludeOriginalData;
		public System.Windows.Forms.TextBox CellEmptyErrorText;
		public System.Windows.Forms.TextBox CellContainsSpacesError;
		public System.Windows.Forms.TextBox CellContainsNonAlphaError;
		public System.Windows.Forms.TextBox CellContainsNumbersError;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label5;
		public System.Windows.Forms.TextBox CellContainsLettersError;
		private System.Windows.Forms.Label label6;
		public System.Windows.Forms.TextBox CellMustBeginWithError;
		private System.Windows.Forms.Label label7;
		public System.Windows.Forms.TextBox CellMustEndWithError;
		private System.Windows.Forms.Label label8;
		public System.Windows.Forms.TextBox CellLengthError;
		private System.Windows.Forms.Label label9;
		public System.Windows.Forms.TextBox CellAllowedItemsError;
		private System.Windows.Forms.Label label13;
		public System.Windows.Forms.TextBox CellNullError;
		private System.Windows.Forms.Label label12;
		public System.Windows.Forms.TextBox CellStringEqualError;
		private System.Windows.Forms.Label label11;
		public System.Windows.Forms.TextBox CellLessThanError;
		private System.Windows.Forms.Label label10;
		public System.Windows.Forms.TextBox CellMoreThanError;
		private System.Windows.Forms.Label label14;
		public System.Windows.Forms.TextBox CellUnexpectedError;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.RadioButton UseColumnNumberRadio;
		private System.Windows.Forms.RadioButton UseHeaderNamesRadio;
		private System.Windows.Forms.ColorDialog colorDialog;
		private System.Windows.Forms.ComboBox ErrorColorCombo;
		private System.Windows.Forms.Label EnumLabel;
		private System.Windows.Forms.Button SaveButton;
		public System.Windows.Forms.CheckBox IncludePivotCheckBox;
		public System.Windows.Forms.CheckBox HyperLinkEnabledCheckbox;
		public System.Windows.Forms.CheckBox NullReportCheckbox;
		public System.Windows.Forms.CheckBox InvalidTypesReportCheckBox;
		private System.Windows.Forms.StatusStrip statusStrip1;
		public System.Windows.Forms.ToolStripStatusLabel RuleSetLocation;
	}
}