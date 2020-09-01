using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Reflection;

namespace WindowsFormsApplication1
{
    // Excel Sheet를 WinForm에서 사용하는 방법 - 두 번째 이야기
    // http://www.sysnet.pe.kr/2/0/12002

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string xlsPath = Path.Combine(Environment.CurrentDirectory, "excel_test.xlsx");
            //this.webBrowser1.Navigate(xlsPath, false);
            this.axFramerControl1.Open(xlsPath);

            //string xls2Path = Path.Combine(Environment.CurrentDirectory, "excel_test2.xlsx");
            // this.axFramerControl1.SaveAs(xls2Path, 12);
        }

        bool _changed = false;

        private void Button1_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = this.axFramerControl1.GetApplication as Microsoft.Office.Interop.Excel.Application;
            if (app == null)
            {
                return;
            }

            Microsoft.Office.Interop.Excel.Worksheet sheet = app.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            if (sheet == null)
            {
                return;
            }

            sheet.Range["A1", "A2"].Cells.Value2 = "test";

            if (_changed == false)
            {
                if (sheet != null)
                {
                    sheet.Change += WorkSheet_Change;
                    _changed = true;
                }
            }
        }

        private void WorkSheet_Change(Microsoft.Office.Interop.Excel.Range target)
        {
            Microsoft.Office.Interop.Excel.Application app = target.Application as Microsoft.Office.Interop.Excel.Application;
            if (app == null)
            {
                return;
            }

            app.EnableEvents = false;

            try
            {
                target.Item[1, 1].EntireRow.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);

                object objValue = target.Item[1, 1].Value2;
                System.Diagnostics.Trace.WriteLine(objValue);
            }
            finally
            {
                app.EnableEvents = true;
            }
        }
    }
}
