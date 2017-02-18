namespace SimpleAddIn
{
	partial class SampleRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// Требуется переменная конструктора.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public SampleRibbon()
			: base(Globals.Factory.GetRibbonFactory())
		{
			InitializeComponent();
		}

		/// <summary> 
		/// Освободить все используемые ресурсы.
		/// </summary>
		/// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null)) {
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Код, автоматически созданный конструктором компонентов

		/// <summary>
		/// Обязательный метод для поддержки конструктора - не изменяйте
		/// содержимое данного метода при помощи редактора кода.
		/// </summary>
		private void InitializeComponent()
		{
			this.tab1 = this.Factory.CreateRibbonTab();
			this.group1 = this.Factory.CreateRibbonGroup();
			this.button1 = this.Factory.CreateRibbonButton();
			this.tab1.SuspendLayout();
			this.group1.SuspendLayout();
			// 
			// tab1
			// 
			this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
			this.tab1.Groups.Add(this.group1);
			this.tab1.Label = "Sample Actions";
			this.tab1.Name = "tab1";
			// 
			// group1
			// 
			this.group1.Items.Add(this.button1);
			this.group1.Label = "General";
			this.group1.Name = "group1";
			// 
			// button1
			// 
			this.button1.Label = "Open JPG";
			this.button1.Name = "button1";
			this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
			// 
			// SampleRibbon
			// 
			this.Name = "SampleRibbon";
			this.RibbonType = "Microsoft.Excel.Workbook";
			this.Tabs.Add(this.tab1);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.SampleRibbon_Load);
			this.tab1.ResumeLayout(false);
			this.tab1.PerformLayout();
			this.group1.ResumeLayout(false);
			this.group1.PerformLayout();

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
	}

	partial class ThisRibbonCollection
	{
		internal SampleRibbon SampleRibbon
		{
			get { return this.GetRibbon<SampleRibbon>(); }
		}
	}
}
