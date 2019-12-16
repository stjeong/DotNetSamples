using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EnumHandles
{
    public partial class Form1 : Form
    {
        FileStream fs;
        public Form1()
        {
            InitializeComponent();

            fs = File.OpenRead("TextFile1.txt");
            this.listView1.Columns.Add("Type", 120);
            this.listView1.Columns.Add("Name", 900);
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            base.OnClosing(e);

            fs.Close();
        }

        private void ListHandles(int processId)
        {
            try
            {
                this.listView1.BeginUpdate();
                this.listView1.Items.Clear();

                using (WindowsHandleInfo whi = new WindowsHandleInfo())
                {
                    for (int i = 0; i < whi.HandleCount; i++)
                    {
                        SYSTEM_HANDLE_ENTRY she = whi[i];

                        if (she.OwnerPid != processId)
                        {
                            continue;
                        }

                        string objName = she.GetName(out string handleTypeName);
                        if (string.IsNullOrEmpty(handleTypeName) == true)
                        {
                            continue;
                        }

                        if (string.IsNullOrEmpty(objName) == true)
                        {
                            continue;
                        }

                        ListViewItem lvItem = this.listView1.Items.Add(handleTypeName);
                        lvItem.SubItems.Add(objName);

                        /*
                        if (File.Exists(objName))
                        {
                            ListViewItem lvItem = this.listView1.Items.Add("File");
                            lvItem.SubItems.Add(objName);
                        }
                        else if (Directory.Exists(objName))
                        {
                            ListViewItem lvItem = this.listView1.Items.Add("Directory");
                            lvItem.SubItems.Add(objName);
                        }
                        */
                    }
                }
            }
            catch (Exception e)
            {
                Trace.WriteLine(e.ToString());
            }

            this.listView1.EndUpdate();
            this.Text = $"# of file handles: {this.listView1.Items.Count}";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.listBox1.Sorted = true;

            foreach (Process process in Process.GetProcesses())
            {
                string txt = $"{process.ProcessName} ({process.Id})";
                this.listBox1.Items.Add(txt);
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string txt = this.listBox1.SelectedItem as string;
            if (txt != null)
            {
                int pid = int.Parse(txt.Substring(txt.IndexOf(" (") + 2).Trim(')'));
                ListHandles(pid);
            }
        }
    }
}
