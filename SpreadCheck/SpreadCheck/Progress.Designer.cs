namespace SpreadCheck
{
	partial class Progress
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
			this.RunProgress = new System.Windows.Forms.ProgressBar();
			this.RowCheckLabel = new System.Windows.Forms.Label();
			this.ColumnCheckLabel = new System.Windows.Forms.Label();
			this.ErrorsFoundLabel = new System.Windows.Forms.Label();
			this.StopButton = new System.Windows.Forms.Button();
			this.ElapsedLabel = new System.Windows.Forms.Label();
			this.FunctionsRunLabel = new System.Windows.Forms.Label();
			this.ProgressLabel = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// RunProgress
			// 
			this.RunProgress.Location = new System.Drawing.Point(30, 186);
			this.RunProgress.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
			this.RunProgress.Name = "RunProgress";
			this.RunProgress.Size = new System.Drawing.Size(756, 51);
			this.RunProgress.TabIndex = 34;
			this.RunProgress.Visible = false;
			// 
			// RowCheckLabel
			// 
			this.RowCheckLabel.AutoSize = true;
			this.RowCheckLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
			this.RowCheckLabel.Location = new System.Drawing.Point(104, 83);
			this.RowCheckLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.RowCheckLabel.Name = "RowCheckLabel";
			this.RowCheckLabel.Size = new System.Drawing.Size(79, 32);
			this.RowCheckLabel.TabIndex = 35;
			this.RowCheckLabel.Text = "Row:";
			// 
			// ColumnCheckLabel
			// 
			this.ColumnCheckLabel.AutoSize = true;
			this.ColumnCheckLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
			this.ColumnCheckLabel.Location = new System.Drawing.Point(62, 46);
			this.ColumnCheckLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.ColumnCheckLabel.Name = "ColumnCheckLabel";
			this.ColumnCheckLabel.Size = new System.Drawing.Size(121, 32);
			this.ColumnCheckLabel.TabIndex = 36;
			this.ColumnCheckLabel.Text = "Column:";
			// 
			// ErrorsFoundLabel
			// 
			this.ErrorsFoundLabel.AutoSize = true;
			this.ErrorsFoundLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
			this.ErrorsFoundLabel.Location = new System.Drawing.Point(372, 46);
			this.ErrorsFoundLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.ErrorsFoundLabel.Name = "ErrorsFoundLabel";
			this.ErrorsFoundLabel.Size = new System.Drawing.Size(187, 32);
			this.ErrorsFoundLabel.TabIndex = 37;
			this.ErrorsFoundLabel.Text = "Errors Found:";
			// 
			// StopButton
			// 
			this.StopButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
			this.StopButton.Location = new System.Drawing.Point(18, 265);
			this.StopButton.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
			this.StopButton.Name = "StopButton";
			this.StopButton.Size = new System.Drawing.Size(768, 51);
			this.StopButton.TabIndex = 38;
			this.StopButton.Text = "STOP!";
			this.StopButton.UseVisualStyleBackColor = true;
			this.StopButton.Click += new System.EventHandler(this.StopButton_Click);
			// 
			// ElapsedLabel
			// 
			this.ElapsedLabel.AutoSize = true;
			this.ElapsedLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
			this.ElapsedLabel.Location = new System.Drawing.Point(36, 155);
			this.ElapsedLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.ElapsedLabel.Name = "ElapsedLabel";
			this.ElapsedLabel.Size = new System.Drawing.Size(138, 25);
			this.ElapsedLabel.TabIndex = 41;
			this.ElapsedLabel.Text = "Time Elapsed:";
			// 
			// FunctionsRunLabel
			// 
			this.FunctionsRunLabel.AutoSize = true;
			this.FunctionsRunLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
			this.FunctionsRunLabel.Location = new System.Drawing.Point(372, 102);
			this.FunctionsRunLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.FunctionsRunLabel.Name = "FunctionsRunLabel";
			this.FunctionsRunLabel.Size = new System.Drawing.Size(206, 32);
			this.FunctionsRunLabel.TabIndex = 42;
			this.FunctionsRunLabel.Text = "Functions Run:";
			// 
			// ProgressLabel
			// 
			this.ProgressLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte) (0)));
			this.ProgressLabel.Location = new System.Drawing.Point(30, 250);
			this.ProgressLabel.Name = "ProgressLabel";
			this.ProgressLabel.Size = new System.Drawing.Size(756, 77);
			this.ProgressLabel.TabIndex = 43;
			this.ProgressLabel.Text = "Sheet Error Checking Complete";
			this.ProgressLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			this.ProgressLabel.Visible = false;
			// 
			// Progress
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(804, 334);
			this.Controls.Add(this.ProgressLabel);
			this.Controls.Add(this.FunctionsRunLabel);
			this.Controls.Add(this.ElapsedLabel);
			this.Controls.Add(this.StopButton);
			this.Controls.Add(this.ErrorsFoundLabel);
			this.Controls.Add(this.ColumnCheckLabel);
			this.Controls.Add(this.RowCheckLabel);
			this.Controls.Add(this.RunProgress);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
			this.Name = "Progress";
			this.Text = "Progress";
			this.ResumeLayout(false);
			this.PerformLayout();
		}

		private System.Windows.Forms.Label ProgressLabel;

		#endregion

		public System.Windows.Forms.Label ElapsedLabel;
		public System.Windows.Forms.Button StopButton;
		public System.Windows.Forms.Label ErrorsFoundLabel;
		public System.Windows.Forms.Label ColumnCheckLabel;
		public System.Windows.Forms.Label RowCheckLabel;
		public System.Windows.Forms.ProgressBar RunProgress;
		public System.Windows.Forms.Label FunctionsRunLabel;
	}
}