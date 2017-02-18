using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace SimpleAddIn
{
	public partial class SampleRibbon
	{
		public event Action ButtonClicked;

		private void SampleRibbon_Load(object sender, RibbonUIEventArgs e)
		{

		}

		private void button1_Click(object sender, RibbonControlEventArgs e)
		{
			if (ButtonClicked != null)
				ButtonClicked();
		}
	}
}
