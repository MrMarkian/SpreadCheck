namespace SpreadCheck
{
	partial class PreviewData
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
			if (disposing && (components != null)) {
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PreviewData));
			this.PreviewList = new System.Windows.Forms.ListBox();
			this.MaxRowNumber = new System.Windows.Forms.NumericUpDown();
			this.RefreshBButton = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.MaxRowNumber)).BeginInit();
			this.SuspendLayout();
			// 
			// PreviewList
			// 
			this.PreviewList.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.PreviewList.FormattingEnabled = true;
			this.PreviewList.Location = new System.Drawing.Point(3, 26);
			this.PreviewList.Name = "PreviewList";
			this.PreviewList.Size = new System.Drawing.Size(404, 420);
			this.PreviewList.TabIndex = 0;
			this.PreviewList.VisibleChanged += new System.EventHandler(this.PreviewList_VisibleChanged);
			this.PreviewList.Resize += new System.EventHandler(this.PreviewList_Resize);
			// 
			// MaxRowNumber
			// 
			this.MaxRowNumber.Dock = System.Windows.Forms.DockStyle.Left;
			this.MaxRowNumber.Location = new System.Drawing.Point(87, 0);
			this.MaxRowNumber.Name = "MaxRowNumber";
			this.MaxRowNumber.Size = new System.Drawing.Size(81, 20);
			this.MaxRowNumber.TabIndex = 2;
			this.MaxRowNumber.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
			// 
			// RefreshBButton
			// 
			this.RefreshBButton.Dock = System.Windows.Forms.DockStyle.Top;
			this.RefreshBButton.Location = new System.Drawing.Point(168, 0);
			this.RefreshBButton.Name = "RefreshBButton";
			this.RefreshBButton.Size = new System.Drawing.Size(239, 23);
			this.RefreshBButton.TabIndex = 3;
			this.RefreshBButton.Text = "Refresh";
			this.RefreshBButton.UseVisualStyleBackColor = true;
			this.RefreshBButton.Click += new System.EventHandler(this.RefreshBButton_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Dock = System.Windows.Forms.DockStyle.Left;
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(0, 0);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(87, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "Return Rows:";
			// 
			// PreviewData
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(407, 447);
			this.Controls.Add(this.RefreshBButton);
			this.Controls.Add(this.MaxRowNumber);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.PreviewList);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "PreviewData";
			this.Text = "Data Preview";
			((System.ComponentModel.ISupportInitialize)(this.MaxRowNumber)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.ListBox PreviewList;
		private System.Windows.Forms.NumericUpDown MaxRowNumber;
		private System.Windows.Forms.Button RefreshBButton;
		private System.Windows.Forms.Label label1;
	}
}