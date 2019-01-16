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
			this.ElapsedLabel = new System.Windows.Forms.Label();
			this.StopButton = new System.Windows.Forms.Button();
			this.SaveExitButton = new System.Windows.Forms.Button();
			this.ErrorsFoundLabel = new System.Windows.Forms.Label();
			this.ColumnCheckLabel = new System.Windows.Forms.Label();
			this.RowCheckLabel = new System.Windows.Forms.Label();
			this.RunProgress = new System.Windows.Forms.ProgressBar();
			this.EmailButton = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// ElapsedLabel
			// 
			this.ElapsedLabel.AutoSize = true;
			this.ElapsedLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.ElapsedLabel.Location = new System.Drawing.Point(31, 108);
			this.ElapsedLabel.Name = "ElapsedLabel";
			this.ElapsedLabel.Size = new System.Drawing.Size(98, 17);
			this.ElapsedLabel.TabIndex = 41;
			this.ElapsedLabel.Text = "Time Elapsed:";
			// 
			// StopButton
			// 
			this.StopButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.StopButton.Location = new System.Drawing.Point(282, 131);
			this.StopButton.Name = "StopButton";
			this.StopButton.Size = new System.Drawing.Size(239, 33);
			this.StopButton.TabIndex = 38;
			this.StopButton.Text = "STOP!";
			this.StopButton.UseVisualStyleBackColor = true;
			this.StopButton.Click += new System.EventHandler(this.StopButton_Click);
			// 
			// SaveExitButton
			// 
			this.SaveExitButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.SaveExitButton.Location = new System.Drawing.Point(398, 134);
			this.SaveExitButton.Name = "SaveExitButton";
			this.SaveExitButton.Size = new System.Drawing.Size(109, 27);
			this.SaveExitButton.TabIndex = 40;
			this.SaveExitButton.Text = "Save Excel File";
			this.SaveExitButton.UseVisualStyleBackColor = true;
			this.SaveExitButton.Visible = false;
			this.SaveExitButton.Click += new System.EventHandler(this.SaveExitButton_Click);
			// 
			// ErrorsFoundLabel
			// 
			this.ErrorsFoundLabel.AutoSize = true;
			this.ErrorsFoundLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.ErrorsFoundLabel.Location = new System.Drawing.Point(265, 36);
			this.ErrorsFoundLabel.Name = "ErrorsFoundLabel";
			this.ErrorsFoundLabel.Size = new System.Drawing.Size(127, 24);
			this.ErrorsFoundLabel.TabIndex = 37;
			this.ErrorsFoundLabel.Text = "Errors Found:";
			// 
			// ColumnCheckLabel
			// 
			this.ColumnCheckLabel.AutoSize = true;
			this.ColumnCheckLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.ColumnCheckLabel.Location = new System.Drawing.Point(41, 30);
			this.ColumnCheckLabel.Name = "ColumnCheckLabel";
			this.ColumnCheckLabel.Size = new System.Drawing.Size(81, 24);
			this.ColumnCheckLabel.TabIndex = 36;
			this.ColumnCheckLabel.Text = "Column:";
			// 
			// RowCheckLabel
			// 
			this.RowCheckLabel.AutoSize = true;
			this.RowCheckLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.RowCheckLabel.Location = new System.Drawing.Point(69, 54);
			this.RowCheckLabel.Name = "RowCheckLabel";
			this.RowCheckLabel.Size = new System.Drawing.Size(53, 24);
			this.RowCheckLabel.TabIndex = 35;
			this.RowCheckLabel.Text = "Row:";
			// 
			// RunProgress
			// 
			this.RunProgress.Location = new System.Drawing.Point(27, 128);
			this.RunProgress.Name = "RunProgress";
			this.RunProgress.Size = new System.Drawing.Size(231, 33);
			this.RunProgress.TabIndex = 34;
			this.RunProgress.Visible = false;
			// 
			// EmailButton
			// 
			this.EmailButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.EmailButton.Location = new System.Drawing.Point(282, 134);
			this.EmailButton.Name = "EmailButton";
			this.EmailButton.Size = new System.Drawing.Size(110, 27);
			this.EmailButton.TabIndex = 39;
			this.EmailButton.Text = "Email Results";
			this.EmailButton.UseVisualStyleBackColor = true;
			this.EmailButton.Visible = false;
			// 
			// Progress
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(537, 216);
			this.Controls.Add(this.ElapsedLabel);
			this.Controls.Add(this.StopButton);
			this.Controls.Add(this.SaveExitButton);
			this.Controls.Add(this.ErrorsFoundLabel);
			this.Controls.Add(this.ColumnCheckLabel);
			this.Controls.Add(this.RowCheckLabel);
			this.Controls.Add(this.RunProgress);
			this.Controls.Add(this.EmailButton);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Name = "Progress";
			this.Text = "Progress";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		public System.Windows.Forms.Label ElapsedLabel;
		public System.Windows.Forms.Button StopButton;
		public System.Windows.Forms.Button SaveExitButton;
		public System.Windows.Forms.Label ErrorsFoundLabel;
		public System.Windows.Forms.Label ColumnCheckLabel;
		public System.Windows.Forms.Label RowCheckLabel;
		public System.Windows.Forms.ProgressBar RunProgress;
		public System.Windows.Forms.Button EmailButton;
	}
}