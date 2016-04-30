using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SKYPE4COMLib;

namespace Magic_Skype_Tool
{
    public partial class Form1 : Form
    {
        Skype skype;
        List<String> contacts = new List<string>();
        AutoCompleteStringCollection searchindex = new AutoCompleteStringCollection();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            System.Threading.Thread thread = new System.Threading.Thread(new System.Threading.ThreadStart(loadContacts));
            thread.IsBackground = true;
            thread.Priority = System.Threading.ThreadPriority.AboveNormal;
            thread.Name = "loadContacts";
            skype = new Skype();
            timer1.Start();
            
        }
        private void loadContacts()
        {
            foreach (User user in skype.Friends)
            {
                contacts.Add(user.Handle);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            skype.Attach();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            int usage = PerfUtils.getCpuUsage();
            progressBar1.Value = usage ;
            label1.Text ="CPU usage:"+ usage.ToString()+"%";
            MessageBox.Show(usage.ToString());
        }
        private void indexContacts()
        {
            foreach(User user in skype.Friends)
            {
                searchindex.Add(user.Handle);
            }
            textBox1.AutoCompleteCustomSource = searchindex;


        }
    }
}
