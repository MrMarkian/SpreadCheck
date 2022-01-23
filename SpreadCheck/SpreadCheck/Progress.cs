using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SpreadCheck
{	public partial class Progress : Form
	{	public Progress()
		{	InitializeComponent();
		}
		private void StopButton_Click(object sender, EventArgs e)
		{	Form1.RulesRunning = false;
		}

	}
}
