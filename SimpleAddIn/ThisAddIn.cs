using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using System.Drawing;
using System.Windows.Forms;
using System.IO;

namespace SimpleAddIn
{
    public partial class ThisAddIn
    {
		private void WriteString(ref MemoryStream stream, string source)
		{
			byte[] buffer = Encoding.UTF8.GetBytes(source);
			stream.Write(buffer, 0, buffer.Length);
		}


		private void GenerateTable(dynamic wSheet, int row, int column)
		{
			OpenFileDialog Dlg = new OpenFileDialog();
			if (Dlg.ShowDialog() == DialogResult.OK) {
				MemoryStream main_stream = new MemoryStream();
				MemoryStream table_stream = new MemoryStream();
				WriteString(ref main_stream, xml_template.head);
				//Application.ScreenUpdating = false;
				Application.Windows[1].DisplayGridlines = false;
				wSheet.Columns.ColumnWidth = 0.08;
				wSheet.Rows.RowHeight = 0.75;
				Bitmap bmp = new Bitmap(Dlg.FileName);
				int index=62;
				for (int j = 0; j < bmp.Height; j++) {
					WriteString(ref table_stream, "   <Row>\r\n");
					for (int i = 0; i < bmp.Width; i++) {
						Color cl = bmp.GetPixel(i, j);
						string style = "  <Style ss:ID=\"s" +
							index.ToString() +
							"\">\r\n   <Interior ss:Color=\"" +
							string.Format("#{0:X2}{1:X2}{2:X2}", cl.R, cl.G, cl.B) +
							"\" ss:Pattern=\"Solid\"/>\r\n  </Style>\r\n";
						WriteString(ref main_stream, style);
						string table_row = "    <Cell ss:StyleID=\"s" + index.ToString() + "\"/>\r\n";
						WriteString(ref table_stream, table_row);
						index++;
						wSheet.Cells[j + 1, i + 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(cl);
					}
					WriteString(ref table_stream, "   </Row>\r\n");
				}
				WriteString(ref main_stream, xml_template.midle);
				string table_head = (bmp.Width+1).ToString() + "\" ss:ExpandedRowCount=\"" + (bmp.Height+1).ToString() + "\">\r\n";
				WriteString(ref main_stream, table_head);
				table_stream.Position=0;
				table_stream.CopyTo(main_stream);
				WriteString(ref main_stream, xml_template.end);
				main_stream.WriteByte(0);
				main_stream.Position = 0;
				//Clipboard.SetData("XML Spreadsheet", main_stream);
				//wSheet.Cells[1, 1].PasteSpecial();
				Application.ScreenUpdating = true;
				/*using (FileStream file = new FileStream("c:\\out.txt",FileMode.OpenOrCreate,FileAccess.ReadWrite)) {
					main_stream.CopyTo(file);
				}

				Application.ScreenUpdating = true;
				/*MemoryStream ms = new MemoryStream();
				using (FileStream file = new FileStream("c:\\out1.txt", FileMode.OpenOrCreate, FileAccess.ReadWrite)) {
					file.CopyTo(ms);
				}
				MessageBox.Show(ms.Length.ToString());
				Clipboard.SetData("XML Spreadsheet", ms);
				wSheet.Cells[1, 1].PasteSpecial();*/
				/*wSheet.Range[wSheet.Cells[1,1],wSheet.Cells[bmp.Height+1,bmp.Width+1]].Copy();
				IDataObject obj = Clipboard.GetDataObject();
				MemoryStream ms= (MemoryStream)obj.GetData("XML Spreadsheet");
				ms.Position = ms.Length - 10;
				while (ms.Position<=ms.Length){
					byte b = (byte)ms.ReadByte();
					MessageBox.Show(b.ToString());
				}
				using (FileStream file = new FileStream("c:\\out1.txt",FileMode.OpenOrCreate,FileAccess.ReadWrite)) {
					ms.CopyTo(file);
				}*/
			}
		}
		

		protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
		{
			var ribbon = new SampleRibbon();
			ribbon.ButtonClicked += ribbon_ButtonClicked;
			return Globals.Factory.GetRibbonFactory().CreateRibbonManager(new IRibbonExtension[] { ribbon });
		}

		private void ribbon_ButtonClicked()
		{
			
			GenerateTable(Application.ActiveWorkbook.ActiveSheet, 1, 1);
		}

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
