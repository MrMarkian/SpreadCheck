namespace SpreadCheck
{
    partial class Form1
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
			this.HeaderList = new System.Windows.Forms.ListBox();
			this.WarnIfEmpty = new System.Windows.Forms.CheckBox();
			this.menuStrip1 = new System.Windows.Forms.MenuStrip();
			this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.loadRuleSetsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.saveRuleSetsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
			this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.optionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.columnToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.columnHeaderStartToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.ColumnHeaderStart = new System.Windows.Forms.ToolStripTextBox();
			this.columnHeaderEndToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.ColumnnHeaderEnd = new System.Windows.Forms.ToolStripTextBox();
			this.rowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.rowHeaderStartToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.RowHeaderStart = new System.Windows.Forms.ToolStripTextBox();
			this.lastRowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.LastRow = new System.Windows.Forms.ToolStripTextBox();
			this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
			this.enableAllRuleSetsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.disableAllRuleSetsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.sheetToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.SheetComboBox = new System.Windows.Forms.ToolStripComboBox();
			this.reportingToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.previewColumnDataToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.RulesGroupBox = new System.Windows.Forms.GroupBox();
			this.MustEndWithCheckBox = new System.Windows.Forms.CheckBox();
			this.button3 = new System.Windows.Forms.Button();
			this.EndWithTextBox = new System.Windows.Forms.TextBox();
			this.GroupBox2 = new System.Windows.Forms.GroupBox();
			this.ItemCheckTextBox = new System.Windows.Forms.TextBox();
			this.Button2 = new System.Windows.Forms.Button();
			this.AllowValuesCheckbox = new System.Windows.Forms.CheckBox();
			this.Button1 = new System.Windows.Forms.Button();
			this.AllowedItemsList = new System.Windows.Forms.ListBox();
			this.CheckDateTimeCheckBox = new System.Windows.Forms.CheckBox();
			this.MustBeginWithCheckbox = new System.Windows.Forms.CheckBox();
			this.LengthEnabledCheckbox = new System.Windows.Forms.CheckBox();
			this.LengthNumericUpDown = new System.Windows.Forms.NumericUpDown();
			this.BeginWithTextBox = new System.Windows.Forms.TextBox();
			this.IfNonAlphaCheckBox = new System.Windows.Forms.CheckBox();
			this.WarnIfContainsSpace = new System.Windows.Forms.CheckBox();
			this.IfNumbersCheckBox = new System.Windows.Forms.CheckBox();
			this.IfLettersCheckBox = new System.Windows.Forms.CheckBox();
			this.ColorCellCheckBox = new System.Windows.Forms.CheckBox();
			this.LessThanNumber = new System.Windows.Forms.NumericUpDown();
			this.MoreThanNumber = new System.Windows.Forms.NumericUpDown();
			this.lessThanCheckBox = new System.Windows.Forms.CheckBox();
			this.MoreThanCheckBox = new System.Windows.Forms.CheckBox();
			this.Functions = new System.Windows.Forms.GroupBox();
			this.NumberFormatCombo = new System.Windows.Forms.ComboBox();
			this.TextAlignmentList = new System.Windows.Forms.ComboBox();
			this.TextRealignCheckbox = new System.Windows.Forms.CheckBox();
			this.ChangeCaseCheckBox = new System.Windows.Forms.CheckBox();
			this.ReverseCheckBox = new System.Windows.Forms.CheckBox();
			this.ChangeCaseCombo = new System.Windows.Forms.ComboBox();
			this.TrimCheckBox = new System.Windows.Forms.CheckBox();
			this.EnabledCheckBox = new System.Windows.Forms.CheckBox();
			this.RunButton = new System.Windows.Forms.Button();
			this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.statusStrip1 = new System.Windows.Forms.StatusStrip();
			this.colrowlabel = new System.Windows.Forms.ToolStripStatusLabel();
			this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
			this.CellFormatLabel = new System.Windows.Forms.ToolStripStatusLabel();
			this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
			this.NumberFormatLabel = new System.Windows.Forms.ToolStripStatusLabel();
			this.StatLabel = new System.Windows.Forms.ToolStripStatusLabel();
			this.StatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
			this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.menuStrip1.SuspendLayout();
			this.RulesGroupBox.SuspendLayout();
			this.GroupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.LengthNumericUpDown)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.LessThanNumber)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.MoreThanNumber)).BeginInit();
			this.Functions.SuspendLayout();
			this.statusStrip1.SuspendLayout();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// HeaderList
			// 
			this.HeaderList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.HeaderList.FormattingEnabled = true;
			this.HeaderList.Location = new System.Drawing.Point(12, 29);
			this.HeaderList.MinimumSize = new System.Drawing.Size(240, 500);
			this.HeaderList.Name = "HeaderList";
			this.HeaderList.Size = new System.Drawing.Size(244, 550);
			this.HeaderList.TabIndex = 0;
			this.HeaderList.SelectedIndexChanged += new System.EventHandler(this.HeaderList_SelectedIndexChanged);
			// 
			// WarnIfEmpty
			// 
			this.WarnIfEmpty.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.WarnIfEmpty.AutoSize = true;
			this.WarnIfEmpty.ForeColor = System.Drawing.SystemColors.WindowText;
			this.WarnIfEmpty.Location = new System.Drawing.Point(14, 26);
			this.WarnIfEmpty.Name = "WarnIfEmpty";
			this.WarnIfEmpty.Size = new System.Drawing.Size(92, 17);
			this.WarnIfEmpty.TabIndex = 1;
			this.WarnIfEmpty.Text = "Warn if Empty";
			this.WarnIfEmpty.UseVisualStyleBackColor = true;
			this.WarnIfEmpty.CheckedChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// menuStrip1
			// 
			this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.optionsToolStripMenuItem,
            this.sheetToolStripMenuItem,
            this.reportingToolStripMenuItem,
            this.previewColumnDataToolStripMenuItem});
			this.menuStrip1.Location = new System.Drawing.Point(0, 0);
			this.menuStrip1.Name = "menuStrip1";
			this.menuStrip1.Size = new System.Drawing.Size(794, 24);
			this.menuStrip1.TabIndex = 2;
			this.menuStrip1.Text = "menuStrip1";
			// 
			// fileToolStripMenuItem
			// 
			this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.toolStripSeparator2,
            this.loadRuleSetsToolStripMenuItem,
            this.saveRuleSetsToolStripMenuItem,
            this.toolStripSeparator3,
            this.exitToolStripMenuItem});
			this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
			this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
			this.fileToolStripMenuItem.Text = "File";
			// 
			// openToolStripMenuItem
			// 
			this.openToolStripMenuItem.Name = "openToolStripMenuItem";
			this.openToolStripMenuItem.Size = new System.Drawing.Size(147, 22);
			this.openToolStripMenuItem.Text = "Open";
			this.openToolStripMenuItem.Click += new System.EventHandler(this.openToolStripMenuItem_Click);
			// 
			// toolStripSeparator2
			// 
			this.toolStripSeparator2.Name = "toolStripSeparator2";
			this.toolStripSeparator2.Size = new System.Drawing.Size(144, 6);
			// 
			// loadRuleSetsToolStripMenuItem
			// 
			this.loadRuleSetsToolStripMenuItem.Name = "loadRuleSetsToolStripMenuItem";
			this.loadRuleSetsToolStripMenuItem.Size = new System.Drawing.Size(147, 22);
			this.loadRuleSetsToolStripMenuItem.Text = "Load RuleSets";
			this.loadRuleSetsToolStripMenuItem.Click += new System.EventHandler(this.loadRuleSetsToolStripMenuItem_Click);
			// 
			// saveRuleSetsToolStripMenuItem
			// 
			this.saveRuleSetsToolStripMenuItem.Name = "saveRuleSetsToolStripMenuItem";
			this.saveRuleSetsToolStripMenuItem.Size = new System.Drawing.Size(147, 22);
			this.saveRuleSetsToolStripMenuItem.Text = "Save RuleSets";
			this.saveRuleSetsToolStripMenuItem.Click += new System.EventHandler(this.saveRuleSetsToolStripMenuItem_Click);
			// 
			// toolStripSeparator3
			// 
			this.toolStripSeparator3.Name = "toolStripSeparator3";
			this.toolStripSeparator3.Size = new System.Drawing.Size(144, 6);
			// 
			// exitToolStripMenuItem
			// 
			this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
			this.exitToolStripMenuItem.Size = new System.Drawing.Size(147, 22);
			this.exitToolStripMenuItem.Text = "Exit";
			this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
			// 
			// optionsToolStripMenuItem
			// 
			this.optionsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.columnToolStripMenuItem,
            this.rowToolStripMenuItem,
            this.toolStripSeparator1,
            this.enableAllRuleSetsToolStripMenuItem,
            this.disableAllRuleSetsToolStripMenuItem});
			this.optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
			this.optionsToolStripMenuItem.Size = new System.Drawing.Size(61, 20);
			this.optionsToolStripMenuItem.Text = "Options";
			// 
			// columnToolStripMenuItem
			// 
			this.columnToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.columnHeaderStartToolStripMenuItem,
            this.columnHeaderEndToolStripMenuItem});
			this.columnToolStripMenuItem.Name = "columnToolStripMenuItem";
			this.columnToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
			this.columnToolStripMenuItem.Text = "Column";
			// 
			// columnHeaderStartToolStripMenuItem
			// 
			this.columnHeaderStartToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ColumnHeaderStart});
			this.columnHeaderStartToolStripMenuItem.Name = "columnHeaderStartToolStripMenuItem";
			this.columnHeaderStartToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
			this.columnHeaderStartToolStripMenuItem.Text = "Column Header Start";
			// 
			// ColumnHeaderStart
			// 
			this.ColumnHeaderStart.Name = "ColumnHeaderStart";
			this.ColumnHeaderStart.Size = new System.Drawing.Size(100, 23);
			this.ColumnHeaderStart.Text = "1";
			// 
			// columnHeaderEndToolStripMenuItem
			// 
			this.columnHeaderEndToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ColumnnHeaderEnd});
			this.columnHeaderEndToolStripMenuItem.Name = "columnHeaderEndToolStripMenuItem";
			this.columnHeaderEndToolStripMenuItem.Size = new System.Drawing.Size(185, 22);
			this.columnHeaderEndToolStripMenuItem.Text = "Column Header End";
			// 
			// ColumnnHeaderEnd
			// 
			this.ColumnnHeaderEnd.Name = "ColumnnHeaderEnd";
			this.ColumnnHeaderEnd.Size = new System.Drawing.Size(100, 23);
			this.ColumnnHeaderEnd.Text = "Auto";
			// 
			// rowToolStripMenuItem
			// 
			this.rowToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.rowHeaderStartToolStripMenuItem,
            this.lastRowToolStripMenuItem});
			this.rowToolStripMenuItem.Name = "rowToolStripMenuItem";
			this.rowToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
			this.rowToolStripMenuItem.Text = "Row";
			// 
			// rowHeaderStartToolStripMenuItem
			// 
			this.rowHeaderStartToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.RowHeaderStart});
			this.rowHeaderStartToolStripMenuItem.Name = "rowHeaderStartToolStripMenuItem";
			this.rowHeaderStartToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
			this.rowHeaderStartToolStripMenuItem.Text = "Row Header Start";
			// 
			// RowHeaderStart
			// 
			this.RowHeaderStart.AccessibleRole = System.Windows.Forms.AccessibleRole.TitleBar;
			this.RowHeaderStart.Name = "RowHeaderStart";
			this.RowHeaderStart.Size = new System.Drawing.Size(100, 23);
			this.RowHeaderStart.Text = "1";
			this.RowHeaderStart.Click += new System.EventHandler(this.RowHeaderStart_Click);
			// 
			// lastRowToolStripMenuItem
			// 
			this.lastRowToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.LastRow});
			this.lastRowToolStripMenuItem.Name = "lastRowToolStripMenuItem";
			this.lastRowToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
			this.lastRowToolStripMenuItem.Text = "Last Row";
			// 
			// LastRow
			// 
			this.LastRow.Name = "LastRow";
			this.LastRow.Size = new System.Drawing.Size(100, 23);
			this.LastRow.Text = "10";
			// 
			// toolStripSeparator1
			// 
			this.toolStripSeparator1.Name = "toolStripSeparator1";
			this.toolStripSeparator1.Size = new System.Drawing.Size(177, 6);
			// 
			// enableAllRuleSetsToolStripMenuItem
			// 
			this.enableAllRuleSetsToolStripMenuItem.Name = "enableAllRuleSetsToolStripMenuItem";
			this.enableAllRuleSetsToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
			this.enableAllRuleSetsToolStripMenuItem.Text = "Enable All RuleSets";
			this.enableAllRuleSetsToolStripMenuItem.Click += new System.EventHandler(this.enableAllRuleSetsToolStripMenuItem_Click);
			// 
			// disableAllRuleSetsToolStripMenuItem
			// 
			this.disableAllRuleSetsToolStripMenuItem.Name = "disableAllRuleSetsToolStripMenuItem";
			this.disableAllRuleSetsToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
			this.disableAllRuleSetsToolStripMenuItem.Text = "Disable All RuleSets";
			this.disableAllRuleSetsToolStripMenuItem.Click += new System.EventHandler(this.disableAllRuleSetsToolStripMenuItem_Click);
			// 
			// sheetToolStripMenuItem
			// 
			this.sheetToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.SheetComboBox});
			this.sheetToolStripMenuItem.Name = "sheetToolStripMenuItem";
			this.sheetToolStripMenuItem.Size = new System.Drawing.Size(48, 20);
			this.sheetToolStripMenuItem.Text = "Sheet";
			// 
			// SheetComboBox
			// 
			this.SheetComboBox.Name = "SheetComboBox";
			this.SheetComboBox.Size = new System.Drawing.Size(121, 23);
			// 
			// reportingToolStripMenuItem
			// 
			this.reportingToolStripMenuItem.Enabled = false;
			this.reportingToolStripMenuItem.Name = "reportingToolStripMenuItem";
			this.reportingToolStripMenuItem.Size = new System.Drawing.Size(71, 20);
			this.reportingToolStripMenuItem.Text = "Reporting";
			this.reportingToolStripMenuItem.Click += new System.EventHandler(this.reportingToolStripMenuItem_Click);
			// 
			// previewColumnDataToolStripMenuItem
			// 
			this.previewColumnDataToolStripMenuItem.Enabled = false;
			this.previewColumnDataToolStripMenuItem.Name = "previewColumnDataToolStripMenuItem";
			this.previewColumnDataToolStripMenuItem.Size = new System.Drawing.Size(133, 20);
			this.previewColumnDataToolStripMenuItem.Text = "Preview Column Data";
			this.previewColumnDataToolStripMenuItem.Click += new System.EventHandler(this.previewColumnDataToolStripMenuItem_Click);
			// 
			// RulesGroupBox
			// 
			this.RulesGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.RulesGroupBox.Controls.Add(this.MustEndWithCheckBox);
			this.RulesGroupBox.Controls.Add(this.button3);
			this.RulesGroupBox.Controls.Add(this.EndWithTextBox);
			this.RulesGroupBox.Controls.Add(this.GroupBox2);
			this.RulesGroupBox.Controls.Add(this.CheckDateTimeCheckBox);
			this.RulesGroupBox.Controls.Add(this.MustBeginWithCheckbox);
			this.RulesGroupBox.Controls.Add(this.LengthEnabledCheckbox);
			this.RulesGroupBox.Controls.Add(this.LengthNumericUpDown);
			this.RulesGroupBox.Controls.Add(this.BeginWithTextBox);
			this.RulesGroupBox.Controls.Add(this.IfNonAlphaCheckBox);
			this.RulesGroupBox.Controls.Add(this.WarnIfContainsSpace);
			this.RulesGroupBox.Controls.Add(this.IfNumbersCheckBox);
			this.RulesGroupBox.Controls.Add(this.WarnIfEmpty);
			this.RulesGroupBox.Controls.Add(this.IfLettersCheckBox);
			this.RulesGroupBox.Enabled = false;
			this.RulesGroupBox.Location = new System.Drawing.Point(276, 61);
			this.RulesGroupBox.MinimumSize = new System.Drawing.Size(320, 480);
			this.RulesGroupBox.Name = "RulesGroupBox";
			this.RulesGroupBox.Size = new System.Drawing.Size(320, 505);
			this.RulesGroupBox.TabIndex = 3;
			this.RulesGroupBox.TabStop = false;
			this.RulesGroupBox.Text = "String - RuleSet";
			this.RulesGroupBox.Enter += new System.EventHandler(this.RulesGroupBox_Enter);
			// 
			// MustEndWithCheckBox
			// 
			this.MustEndWithCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.MustEndWithCheckBox.AutoSize = true;
			this.MustEndWithCheckBox.Location = new System.Drawing.Point(14, 294);
			this.MustEndWithCheckBox.Name = "MustEndWithCheckBox";
			this.MustEndWithCheckBox.Size = new System.Drawing.Size(96, 17);
			this.MustEndWithCheckBox.TabIndex = 22;
			this.MustEndWithCheckBox.Text = "Must End With";
			this.MustEndWithCheckBox.UseVisualStyleBackColor = true;
			this.MustEndWithCheckBox.CheckedChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// button3
			// 
			this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.button3.Location = new System.Drawing.Point(149, 174);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(136, 22);
			this.button3.TabIndex = 22;
			this.button3.Text = "DT Settings";
			this.button3.UseVisualStyleBackColor = true;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// EndWithTextBox
			// 
			this.EndWithTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.EndWithTextBox.Location = new System.Drawing.Point(149, 292);
			this.EndWithTextBox.Name = "EndWithTextBox";
			this.EndWithTextBox.Size = new System.Drawing.Size(165, 20);
			this.EndWithTextBox.TabIndex = 21;
			this.EndWithTextBox.TextChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// GroupBox2
			// 
			this.GroupBox2.Controls.Add(this.ItemCheckTextBox);
			this.GroupBox2.Controls.Add(this.Button2);
			this.GroupBox2.Controls.Add(this.AllowValuesCheckbox);
			this.GroupBox2.Controls.Add(this.Button1);
			this.GroupBox2.Controls.Add(this.AllowedItemsList);
			this.GroupBox2.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.GroupBox2.Location = new System.Drawing.Point(3, 350);
			this.GroupBox2.Name = "GroupBox2";
			this.GroupBox2.Size = new System.Drawing.Size(314, 152);
			this.GroupBox2.TabIndex = 20;
			this.GroupBox2.TabStop = false;
			// 
			// ItemCheckTextBox
			// 
			this.ItemCheckTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
			this.ItemCheckTextBox.Location = new System.Drawing.Point(11, 48);
			this.ItemCheckTextBox.Name = "ItemCheckTextBox";
			this.ItemCheckTextBox.Size = new System.Drawing.Size(111, 20);
			this.ItemCheckTextBox.TabIndex = 7;
			this.ItemCheckTextBox.TextChanged += new System.EventHandler(this.ItemCheckTextBox_TextChanged);
			// 
			// Button2
			// 
			this.Button2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
			this.Button2.Location = new System.Drawing.Point(29, 112);
			this.Button2.Name = "Button2";
			this.Button2.Size = new System.Drawing.Size(111, 28);
			this.Button2.TabIndex = 6;
			this.Button2.Text = "Remove Item";
			this.Button2.UseVisualStyleBackColor = true;
			this.Button2.Click += new System.EventHandler(this.Button2_Click_1);
			// 
			// AllowValuesCheckbox
			// 
			this.AllowValuesCheckbox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
			this.AllowValuesCheckbox.AutoSize = true;
			this.AllowValuesCheckbox.Location = new System.Drawing.Point(11, 25);
			this.AllowValuesCheckbox.Name = "AllowValuesCheckbox";
			this.AllowValuesCheckbox.Size = new System.Drawing.Size(112, 17);
			this.AllowValuesCheckbox.TabIndex = 19;
			this.AllowValuesCheckbox.Text = "Only Allow this list:";
			this.AllowValuesCheckbox.UseVisualStyleBackColor = true;
			this.AllowValuesCheckbox.CheckedChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// Button1
			// 
			this.Button1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
			this.Button1.Location = new System.Drawing.Point(29, 83);
			this.Button1.Name = "Button1";
			this.Button1.Size = new System.Drawing.Size(111, 23);
			this.Button1.TabIndex = 5;
			this.Button1.Text = "Add Item";
			this.Button1.UseVisualStyleBackColor = true;
			this.Button1.Click += new System.EventHandler(this.Button1_Click);
			// 
			// AllowedItemsList
			// 
			this.AllowedItemsList.Dock = System.Windows.Forms.DockStyle.Right;
			this.AllowedItemsList.FormattingEnabled = true;
			this.AllowedItemsList.Location = new System.Drawing.Point(165, 16);
			this.AllowedItemsList.Name = "AllowedItemsList";
			this.AllowedItemsList.Size = new System.Drawing.Size(146, 133);
			this.AllowedItemsList.TabIndex = 3;
			// 
			// CheckDateTimeCheckBox
			// 
			this.CheckDateTimeCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
			this.CheckDateTimeCheckBox.AutoSize = true;
			this.CheckDateTimeCheckBox.Location = new System.Drawing.Point(14, 179);
			this.CheckDateTimeCheckBox.Name = "CheckDateTimeCheckBox";
			this.CheckDateTimeCheckBox.Size = new System.Drawing.Size(117, 17);
			this.CheckDateTimeCheckBox.TabIndex = 21;
			this.CheckDateTimeCheckBox.Text = "Check Date / Time";
			this.CheckDateTimeCheckBox.UseVisualStyleBackColor = true;
			// 
			// MustBeginWithCheckbox
			// 
			this.MustBeginWithCheckbox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.MustBeginWithCheckbox.AutoSize = true;
			this.MustBeginWithCheckbox.Location = new System.Drawing.Point(14, 271);
			this.MustBeginWithCheckbox.Name = "MustBeginWithCheckbox";
			this.MustBeginWithCheckbox.Size = new System.Drawing.Size(104, 17);
			this.MustBeginWithCheckbox.TabIndex = 18;
			this.MustBeginWithCheckbox.Text = "Must Begin With";
			this.MustBeginWithCheckbox.UseVisualStyleBackColor = true;
			this.MustBeginWithCheckbox.CheckedChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// LengthEnabledCheckbox
			// 
			this.LengthEnabledCheckbox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.LengthEnabledCheckbox.AutoSize = true;
			this.LengthEnabledCheckbox.ForeColor = System.Drawing.SystemColors.WindowText;
			this.LengthEnabledCheckbox.Location = new System.Drawing.Point(14, 317);
			this.LengthEnabledCheckbox.Name = "LengthEnabledCheckbox";
			this.LengthEnabledCheckbox.Size = new System.Drawing.Size(129, 17);
			this.LengthEnabledCheckbox.TabIndex = 17;
			this.LengthEnabledCheckbox.Text = "Must have a length of";
			this.LengthEnabledCheckbox.UseVisualStyleBackColor = true;
			this.LengthEnabledCheckbox.CheckedChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// LengthNumericUpDown
			// 
			this.LengthNumericUpDown.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.LengthNumericUpDown.Enabled = false;
			this.LengthNumericUpDown.Location = new System.Drawing.Point(149, 317);
			this.LengthNumericUpDown.Name = "LengthNumericUpDown";
			this.LengthNumericUpDown.Size = new System.Drawing.Size(165, 20);
			this.LengthNumericUpDown.TabIndex = 16;
			this.LengthNumericUpDown.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.LengthNumericUpDown.ValueChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// BeginWithTextBox
			// 
			this.BeginWithTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.BeginWithTextBox.Location = new System.Drawing.Point(149, 269);
			this.BeginWithTextBox.Name = "BeginWithTextBox";
			this.BeginWithTextBox.Size = new System.Drawing.Size(165, 20);
			this.BeginWithTextBox.TabIndex = 15;
			this.BeginWithTextBox.TextChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// IfNonAlphaCheckBox
			// 
			this.IfNonAlphaCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.IfNonAlphaCheckBox.AutoSize = true;
			this.IfNonAlphaCheckBox.Location = new System.Drawing.Point(14, 72);
			this.IfNonAlphaCheckBox.Name = "IfNonAlphaCheckBox";
			this.IfNonAlphaCheckBox.Size = new System.Drawing.Size(195, 17);
			this.IfNonAlphaCheckBox.TabIndex = 14;
			this.IfNonAlphaCheckBox.Text = "Warn If Contains Non-Alphanumeric";
			this.IfNonAlphaCheckBox.UseVisualStyleBackColor = true;
			this.IfNonAlphaCheckBox.CheckedChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// WarnIfContainsSpace
			// 
			this.WarnIfContainsSpace.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.WarnIfContainsSpace.AutoSize = true;
			this.WarnIfContainsSpace.Location = new System.Drawing.Point(14, 49);
			this.WarnIfContainsSpace.Name = "WarnIfContainsSpace";
			this.WarnIfContainsSpace.Size = new System.Drawing.Size(143, 17);
			this.WarnIfContainsSpace.TabIndex = 2;
			this.WarnIfContainsSpace.Text = "Warn if Contains Spaces";
			this.WarnIfContainsSpace.UseVisualStyleBackColor = true;
			this.WarnIfContainsSpace.CheckedChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// IfNumbersCheckBox
			// 
			this.IfNumbersCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.IfNumbersCheckBox.AutoSize = true;
			this.IfNumbersCheckBox.ForeColor = System.Drawing.SystemColors.WindowText;
			this.IfNumbersCheckBox.Location = new System.Drawing.Point(14, 95);
			this.IfNumbersCheckBox.Name = "IfNumbersCheckBox";
			this.IfNumbersCheckBox.Size = new System.Drawing.Size(150, 17);
			this.IfNumbersCheckBox.TabIndex = 6;
			this.IfNumbersCheckBox.Text = "Warn If Contains Numbers";
			this.IfNumbersCheckBox.UseVisualStyleBackColor = true;
			this.IfNumbersCheckBox.CheckedChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// IfLettersCheckBox
			// 
			this.IfLettersCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.IfLettersCheckBox.AutoSize = true;
			this.IfLettersCheckBox.Location = new System.Drawing.Point(14, 118);
			this.IfLettersCheckBox.Name = "IfLettersCheckBox";
			this.IfLettersCheckBox.Size = new System.Drawing.Size(140, 17);
			this.IfLettersCheckBox.TabIndex = 13;
			this.IfLettersCheckBox.Text = "Warn If Contains Letters";
			this.IfLettersCheckBox.UseVisualStyleBackColor = true;
			this.IfLettersCheckBox.CheckedChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// ColorCellCheckBox
			// 
			this.ColorCellCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.ColorCellCheckBox.AutoSize = true;
			this.ColorCellCheckBox.Location = new System.Drawing.Point(419, 38);
			this.ColorCellCheckBox.Name = "ColorCellCheckBox";
			this.ColorCellCheckBox.Size = new System.Drawing.Size(135, 17);
			this.ColorCellCheckBox.TabIndex = 28;
			this.ColorCellCheckBox.Text = "BackColor Cell on Error";
			this.ColorCellCheckBox.UseVisualStyleBackColor = true;
			this.ColorCellCheckBox.CheckedChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// LessThanNumber
			// 
			this.LessThanNumber.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.LessThanNumber.DecimalPlaces = 6;
			this.LessThanNumber.Enabled = false;
			this.LessThanNumber.Location = new System.Drawing.Point(8, 104);
			this.LessThanNumber.Maximum = new decimal(new int[] {
            100000000,
            0,
            0,
            0});
			this.LessThanNumber.Minimum = new decimal(new int[] {
            100000000,
            0,
            0,
            -2147483648});
			this.LessThanNumber.Name = "LessThanNumber";
			this.LessThanNumber.Size = new System.Drawing.Size(165, 20);
			this.LessThanNumber.TabIndex = 27;
			this.LessThanNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.LessThanNumber.ValueChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// MoreThanNumber
			// 
			this.MoreThanNumber.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.MoreThanNumber.DecimalPlaces = 6;
			this.MoreThanNumber.Enabled = false;
			this.MoreThanNumber.Location = new System.Drawing.Point(8, 44);
			this.MoreThanNumber.Maximum = new decimal(new int[] {
            100000000,
            0,
            0,
            0});
			this.MoreThanNumber.Minimum = new decimal(new int[] {
            100000000,
            0,
            0,
            -2147483648});
			this.MoreThanNumber.Name = "MoreThanNumber";
			this.MoreThanNumber.Size = new System.Drawing.Size(165, 20);
			this.MoreThanNumber.TabIndex = 23;
			this.MoreThanNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.MoreThanNumber.ValueChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// lessThanCheckBox
			// 
			this.lessThanCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.lessThanCheckBox.AutoSize = true;
			this.lessThanCheckBox.Location = new System.Drawing.Point(11, 81);
			this.lessThanCheckBox.Name = "lessThanCheckBox";
			this.lessThanCheckBox.Size = new System.Drawing.Size(116, 17);
			this.lessThanCheckBox.TabIndex = 26;
			this.lessThanCheckBox.Text = "Warn if Less Than:";
			this.lessThanCheckBox.UseVisualStyleBackColor = true;
			this.lessThanCheckBox.CheckedChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// MoreThanCheckBox
			// 
			this.MoreThanCheckBox.AutoSize = true;
			this.MoreThanCheckBox.Location = new System.Drawing.Point(9, 21);
			this.MoreThanCheckBox.Name = "MoreThanCheckBox";
			this.MoreThanCheckBox.Size = new System.Drawing.Size(118, 17);
			this.MoreThanCheckBox.TabIndex = 25;
			this.MoreThanCheckBox.Text = "Warn if More Than:";
			this.MoreThanCheckBox.UseVisualStyleBackColor = true;
			this.MoreThanCheckBox.CheckedChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// Functions
			// 
			this.Functions.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.Functions.Controls.Add(this.NumberFormatCombo);
			this.Functions.Controls.Add(this.TextAlignmentList);
			this.Functions.Controls.Add(this.TextRealignCheckbox);
			this.Functions.Controls.Add(this.ChangeCaseCheckBox);
			this.Functions.Controls.Add(this.ReverseCheckBox);
			this.Functions.Controls.Add(this.ChangeCaseCombo);
			this.Functions.Controls.Add(this.TrimCheckBox);
			this.Functions.Location = new System.Drawing.Point(602, 276);
			this.Functions.MaximumSize = new System.Drawing.Size(180, 400);
			this.Functions.MinimumSize = new System.Drawing.Size(150, 220);
			this.Functions.Name = "Functions";
			this.Functions.Size = new System.Drawing.Size(180, 284);
			this.Functions.TabIndex = 23;
			this.Functions.TabStop = false;
			this.Functions.Text = "Functions";
			// 
			// NumberFormatCombo
			// 
			this.NumberFormatCombo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.NumberFormatCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.NumberFormatCombo.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.NumberFormatCombo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.NumberFormatCombo.FormattingEnabled = true;
			this.NumberFormatCombo.Location = new System.Drawing.Point(6, 256);
			this.NumberFormatCombo.Name = "NumberFormatCombo";
			this.NumberFormatCombo.Size = new System.Drawing.Size(160, 21);
			this.NumberFormatCombo.TabIndex = 9;
			this.NumberFormatCombo.SelectedIndexChanged += new System.EventHandler(this.NumberFormatCombo_SelectedIndexChanged);
			// 
			// TextAlignmentList
			// 
			this.TextAlignmentList.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.TextAlignmentList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.TextAlignmentList.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.TextAlignmentList.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.TextAlignmentList.FormattingEnabled = true;
			this.TextAlignmentList.Items.AddRange(new object[] {
            "No Change of Alignment",
            "Left",
            "Right",
            "Center"});
			this.TextAlignmentList.Location = new System.Drawing.Point(6, 216);
			this.TextAlignmentList.Name = "TextAlignmentList";
			this.TextAlignmentList.Size = new System.Drawing.Size(160, 21);
			this.TextAlignmentList.TabIndex = 8;
			this.TextAlignmentList.SelectedIndexChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// TextRealignCheckbox
			// 
			this.TextRealignCheckbox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.TextRealignCheckbox.AutoSize = true;
			this.TextRealignCheckbox.Location = new System.Drawing.Point(6, 193);
			this.TextRealignCheckbox.Name = "TextRealignCheckbox";
			this.TextRealignCheckbox.Size = new System.Drawing.Size(73, 17);
			this.TextRealignCheckbox.TabIndex = 7;
			this.TextRealignCheckbox.Text = "Align Text";
			this.TextRealignCheckbox.UseVisualStyleBackColor = true;
			this.TextRealignCheckbox.CheckedChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// ChangeCaseCheckBox
			// 
			this.ChangeCaseCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
			this.ChangeCaseCheckBox.AutoSize = true;
			this.ChangeCaseCheckBox.Location = new System.Drawing.Point(6, 88);
			this.ChangeCaseCheckBox.Name = "ChangeCaseCheckBox";
			this.ChangeCaseCheckBox.Size = new System.Drawing.Size(90, 17);
			this.ChangeCaseCheckBox.TabIndex = 6;
			this.ChangeCaseCheckBox.Text = "Change Case";
			this.ChangeCaseCheckBox.UseVisualStyleBackColor = true;
			this.ChangeCaseCheckBox.CheckedChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// ReverseCheckBox
			// 
			this.ReverseCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.ReverseCheckBox.AutoSize = true;
			this.ReverseCheckBox.Location = new System.Drawing.Point(6, 54);
			this.ReverseCheckBox.Name = "ReverseCheckBox";
			this.ReverseCheckBox.Size = new System.Drawing.Size(92, 17);
			this.ReverseCheckBox.TabIndex = 5;
			this.ReverseCheckBox.Text = "Reverse Data";
			this.ReverseCheckBox.UseVisualStyleBackColor = true;
			this.ReverseCheckBox.CheckedChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// ChangeCaseCombo
			// 
			this.ChangeCaseCombo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
			this.ChangeCaseCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.ChangeCaseCombo.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
			this.ChangeCaseCombo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.ChangeCaseCombo.FormattingEnabled = true;
			this.ChangeCaseCombo.Items.AddRange(new object[] {
            "No Change of Case",
            "To Upper Case",
            "To Lower Case",
            "To Upper Invariant",
            "To Lower Invariant"});
			this.ChangeCaseCombo.Location = new System.Drawing.Point(6, 111);
			this.ChangeCaseCombo.Name = "ChangeCaseCombo";
			this.ChangeCaseCombo.Size = new System.Drawing.Size(160, 21);
			this.ChangeCaseCombo.TabIndex = 3;
			this.ChangeCaseCombo.SelectedIndexChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// TrimCheckBox
			// 
			this.TrimCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.TrimCheckBox.AutoSize = true;
			this.TrimCheckBox.Location = new System.Drawing.Point(6, 19);
			this.TrimCheckBox.Name = "TrimCheckBox";
			this.TrimCheckBox.Size = new System.Drawing.Size(66, 17);
			this.TrimCheckBox.TabIndex = 2;
			this.TrimCheckBox.Text = "Trim Cell";
			this.TrimCheckBox.UseVisualStyleBackColor = true;
			this.TrimCheckBox.CheckedChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// EnabledCheckBox
			// 
			this.EnabledCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.EnabledCheckBox.AutoSize = true;
			this.EnabledCheckBox.Enabled = false;
			this.EnabledCheckBox.Location = new System.Drawing.Point(286, 38);
			this.EnabledCheckBox.Name = "EnabledCheckBox";
			this.EnabledCheckBox.Size = new System.Drawing.Size(65, 17);
			this.EnabledCheckBox.TabIndex = 22;
			this.EnabledCheckBox.Text = "Enabled";
			this.EnabledCheckBox.UseVisualStyleBackColor = true;
			this.EnabledCheckBox.CheckedChanged += new System.EventHandler(this.CheckedChanged);
			// 
			// RunButton
			// 
			this.RunButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.RunButton.Enabled = false;
			this.RunButton.Location = new System.Drawing.Point(578, 31);
			this.RunButton.Name = "RunButton";
			this.RunButton.Size = new System.Drawing.Size(202, 29);
			this.RunButton.TabIndex = 21;
			this.RunButton.Text = "Run!";
			this.RunButton.UseVisualStyleBackColor = true;
			this.RunButton.Click += new System.EventHandler(this.RunButton_Click);
			// 
			// openFileDialog
			// 
			this.openFileDialog.FileName = "*.xlsx";
			// 
			// statusStrip1
			// 
			this.statusStrip1.AllowItemReorder = true;
			this.statusStrip1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.statusStrip1.Dock = System.Windows.Forms.DockStyle.None;
			this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.colrowlabel,
            this.toolStripStatusLabel1,
            this.CellFormatLabel,
            this.toolStripStatusLabel2,
            this.NumberFormatLabel,
            this.StatLabel,
            this.StatusLabel});
			this.statusStrip1.Location = new System.Drawing.Point(476, 569);
			this.statusStrip1.Name = "statusStrip1";
			this.statusStrip1.Size = new System.Drawing.Size(318, 22);
			this.statusStrip1.SizingGrip = false;
			this.statusStrip1.TabIndex = 24;
			this.statusStrip1.Text = "statusStrip1";
			// 
			// colrowlabel
			// 
			this.colrowlabel.Name = "colrowlabel";
			this.colrowlabel.Size = new System.Drawing.Size(31, 17);
			this.colrowlabel.Text = "C: R:";
			// 
			// toolStripStatusLabel1
			// 
			this.toolStripStatusLabel1.Font = new System.Drawing.Font("Segoe UI", 9F);
			this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
			this.toolStripStatusLabel1.Size = new System.Drawing.Size(71, 17);
			this.toolStripStatusLabel1.Text = "CellFormat: ";
			// 
			// CellFormatLabel
			// 
			this.CellFormatLabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
			this.CellFormatLabel.Name = "CellFormatLabel";
			this.CellFormatLabel.Size = new System.Drawing.Size(23, 17);
			this.CellFormatLabel.Text = "<>";
			// 
			// toolStripStatusLabel2
			// 
			this.toolStripStatusLabel2.Name = "toolStripStatusLabel2";
			this.toolStripStatusLabel2.Size = new System.Drawing.Size(95, 17);
			this.toolStripStatusLabel2.Text = "NumberFormat: ";
			// 
			// NumberFormatLabel
			// 
			this.NumberFormatLabel.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
			this.NumberFormatLabel.Name = "NumberFormatLabel";
			this.NumberFormatLabel.Size = new System.Drawing.Size(23, 17);
			this.NumberFormatLabel.Text = "<>";
			// 
			// StatLabel
			// 
			this.StatLabel.Name = "StatLabel";
			this.StatLabel.Size = new System.Drawing.Size(23, 17);
			this.StatLabel.Text = "<>";
			// 
			// StatusLabel
			// 
			this.StatusLabel.Name = "StatusLabel";
			this.StatusLabel.Size = new System.Drawing.Size(35, 17);
			this.StatusLabel.Text = "Idle...";
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.groupBox1.Controls.Add(this.MoreThanCheckBox);
			this.groupBox1.Controls.Add(this.lessThanCheckBox);
			this.groupBox1.Controls.Add(this.LessThanNumber);
			this.groupBox1.Controls.Add(this.MoreThanNumber);
			this.groupBox1.Location = new System.Drawing.Point(602, 66);
			this.groupBox1.MaximumSize = new System.Drawing.Size(200, 200);
			this.groupBox1.MinimumSize = new System.Drawing.Size(180, 140);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(180, 140);
			this.groupBox1.TabIndex = 28;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Double - RuleSet";
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(794, 591);
			this.Controls.Add(this.ColorCellCheckBox);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.statusStrip1);
			this.Controls.Add(this.RunButton);
			this.Controls.Add(this.EnabledCheckBox);
			this.Controls.Add(this.RulesGroupBox);
			this.Controls.Add(this.Functions);
			this.Controls.Add(this.HeaderList);
			this.Controls.Add(this.menuStrip1);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MainMenuStrip = this.menuStrip1;
			this.MinimumSize = new System.Drawing.Size(810, 630);
			this.Name = "Form1";
			this.Text = "Sheet Checker";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
			this.Load += new System.EventHandler(this.Form1_Load);
			this.menuStrip1.ResumeLayout(false);
			this.menuStrip1.PerformLayout();
			this.RulesGroupBox.ResumeLayout(false);
			this.RulesGroupBox.PerformLayout();
			this.GroupBox2.ResumeLayout(false);
			this.GroupBox2.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.LengthNumericUpDown)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.LessThanNumber)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.MoreThanNumber)).EndInit();
			this.Functions.ResumeLayout(false);
			this.Functions.PerformLayout();
			this.statusStrip1.ResumeLayout(false);
			this.statusStrip1.PerformLayout();
			this.groupBox1.ResumeLayout(false);
			this.groupBox1.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox HeaderList;
        private System.Windows.Forms.CheckBox WarnIfEmpty;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem optionsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem columnToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem columnHeaderStartToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem columnHeaderEndToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem rowToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem rowHeaderStartToolStripMenuItem;
        private System.Windows.Forms.ToolStripTextBox RowHeaderStart;
        private System.Windows.Forms.ToolStripMenuItem lastRowToolStripMenuItem;
        private System.Windows.Forms.GroupBox RulesGroupBox;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        internal System.Windows.Forms.CheckBox EnabledCheckBox;
        internal System.Windows.Forms.Button RunButton;
        internal System.Windows.Forms.CheckBox AllowValuesCheckbox;
        internal System.Windows.Forms.GroupBox GroupBox2;
        internal System.Windows.Forms.TextBox ItemCheckTextBox;
        internal System.Windows.Forms.Button Button2;
        internal System.Windows.Forms.Button Button1;
        internal System.Windows.Forms.CheckBox MustBeginWithCheckbox;
        internal System.Windows.Forms.CheckBox LengthEnabledCheckbox;
        internal System.Windows.Forms.NumericUpDown LengthNumericUpDown;
        internal System.Windows.Forms.TextBox BeginWithTextBox;
        internal System.Windows.Forms.CheckBox IfNonAlphaCheckBox;
        internal System.Windows.Forms.CheckBox IfLettersCheckBox;
        internal System.Windows.Forms.CheckBox IfNumbersCheckBox;
        private System.Windows.Forms.CheckBox WarnIfContainsSpace;
        private System.Windows.Forms.ToolStripMenuItem sheetToolStripMenuItem;
        private System.Windows.Forms.ToolStripComboBox SheetComboBox;
        private System.Windows.Forms.ToolStripTextBox LastRow;
        public System.Windows.Forms.ToolStripTextBox ColumnHeaderStart;
        public System.Windows.Forms.ToolStripTextBox ColumnnHeaderEnd;
        private System.Windows.Forms.GroupBox Functions;
        private System.Windows.Forms.CheckBox TrimCheckBox;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel colrowlabel;
        private System.Windows.Forms.ToolStripStatusLabel StatusLabel;
        private System.Windows.Forms.CheckBox CheckDateTimeCheckBox;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.ComboBox ChangeCaseCombo;
        internal System.Windows.Forms.CheckBox MustEndWithCheckBox;
        internal System.Windows.Forms.TextBox EndWithTextBox;
        internal System.Windows.Forms.ListBox AllowedItemsList;
        private System.Windows.Forms.CheckBox ReverseCheckBox;
        private System.Windows.Forms.ToolStripMenuItem reportingToolStripMenuItem;
        private System.Windows.Forms.CheckBox MoreThanCheckBox;
        private System.Windows.Forms.CheckBox lessThanCheckBox;
        internal System.Windows.Forms.NumericUpDown MoreThanNumber;
        internal System.Windows.Forms.NumericUpDown LessThanNumber;
        private System.Windows.Forms.CheckBox ChangeCaseCheckBox;
        internal System.Windows.Forms.CheckBox ColorCellCheckBox;
		private System.Windows.Forms.CheckBox TextRealignCheckbox;
		private System.Windows.Forms.ComboBox TextAlignmentList;
		private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
		private System.Windows.Forms.ToolStripStatusLabel CellFormatLabel;
		private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
		private System.Windows.Forms.ToolStripStatusLabel NumberFormatLabel;
		private System.Windows.Forms.ComboBox NumberFormatCombo;
		private System.Windows.Forms.ToolStripStatusLabel StatLabel;
		private System.Windows.Forms.ToolStripMenuItem enableAllRuleSetsToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem disableAllRuleSetsToolStripMenuItem;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
		private System.Windows.Forms.ToolStripMenuItem saveRuleSetsToolStripMenuItem;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
		private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
		private System.Windows.Forms.ToolStripMenuItem loadRuleSetsToolStripMenuItem;
		private System.Windows.Forms.SaveFileDialog saveFileDialog;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.ToolStripMenuItem previewColumnDataToolStripMenuItem;
	}
}

