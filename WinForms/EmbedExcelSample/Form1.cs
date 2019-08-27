using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

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

        private void Button1_Click(object sender, EventArgs e)
        {
            dynamic doc = this.axFramerControl1.ActiveDocument;
            doc.ActiveSheet.Range("A1:A2").Cells.value = "test";
        }
    }
}
