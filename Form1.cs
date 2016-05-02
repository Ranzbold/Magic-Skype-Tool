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
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Magic_Skype_Tool
{
    public partial class Form1 : Form
    {
        Skype skype;
        List<String> contacts = new List<string>();
        int amount;
        String message;
        AutoCompleteStringCollection searchindex = new AutoCompleteStringCollection();
        Stopwatch stopwatch = new Stopwatch();
        String version = "0.1";






        public Form1()
        {
            InitializeComponent();
        }
        public static void EnableTab(TabPage page, bool enable)
        {
            foreach (Control ctl in page.Controls) ctl.Enabled = enable;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            foreach (TabPage page in tabControl1.TabPages)
            {
                EnableTab(page, false);
            }
            skype = new Skype();
            EnableTab(tabPage1, true);
            ((_ISkypeEvents_Event)skype).AttachmentStatus += OnAttach;
            skype.MessageStatus += OnMessage;
            textBox17.Text = DateTime.Now.ToString();
            timer3.Start();
        }
        private void OnAttach(TAttachmentStatus status)
        {

            if (status.Equals(TAttachmentStatus.apiAttachSuccess))
            {
                MessageBox.Show("Successfully Hooked into Skype");
                backgroundWorker1.WorkerSupportsCancellation = true;
                backgroundWorker1.WorkerReportsProgress = true;
                foreach (TabPage page in tabControl1.TabPages)
                {
                    EnableTab(page, true);
                }
            }
            if (status.Equals(TAttachmentStatus.apiAttachRefused))
            {
                MessageBox.Show("Could not hook into Skype");
            }


        }
        private void OnMessage(ChatMessage msg, TChatMessageStatus status)
        {
            if (status == TChatMessageStatus.cmsReceived)
            {

                skype.SendMessage(msg.Sender.Handle, "TEST");
            }

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
            try
            {
                skype.Attach(7, false);
                loadContacts();
                indexContacts();
            }
            catch (COMException)
            {
                MessageBox.Show("Attach Request could not be sent. Try to restart Skype and the Application!");
            }




        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            int usage = PerfUtils.getCpuUsage();
        }
        private void indexContacts()
        {
            foreach (User user in skype.Friends)
            {
                searchindex.Add(user.Handle);
            }
            textBox1.AutoCompleteCustomSource = searchindex;
            textBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBox1.AutoCompleteSource = AutoCompleteSource.CustomSource;
            textBox2.AutoCompleteCustomSource = searchindex;
            textBox2.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            textBox2.AutoCompleteSource = AutoCompleteSource.CustomSource;
            foreach (var user in contacts)
            {
                checkedListBox1.Items.Add(user);
            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox1.Items.Count - 1; i++)
            {
                if (checkedListBox1.Items[i].ToString().ToLower().Contains(textBox1.Text))
                {
                    checkedListBox1.SetSelected(i, true);
                    checkedListBox1.SelectedIndex = i;
                }


            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int i = 0;
            while (i < checkedListBox1.CheckedItems.Count)
            {
                skype.SendMessage(checkedListBox1.CheckedItems[i].ToString(), richTextBox1.Text);
                i++;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            timer2.Start();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            timer2.Stop();
        }

        private void trackBar1_ValueChanged(object sender, EventArgs e)
        {
            label3.Text = "Speed: " + trackBar1.Value + "ms";
            timer2.Interval = trackBar1.Value;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            foreach (User user in skype.Friends)
            {
                skype.SendMessage(user.Handle, richTextBox1.Text);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (!(richTextBox1.Text == ""))
            {
                String s = Microsoft.VisualBasic.Interaction.InputBox("Define the amount of messages to send", "Amount of messages");

                if (int.TryParse(s, out amount))
                {
                    if (amount > 0)
                    {
                        stopwatch.Start();
                        message = richTextBox1.Text;
                        backgroundWorker1.RunWorkerAsync();
                    }
                    else
                    {
                        MessageBox.Show("The Input has to be greater than 0");
                    }

                }
                else
                {
                    MessageBox.Show("The Input has to be an even number");
                }
            }
            else
            {
                MessageBox.Show("The text to send must not be empty");
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            sendMessages();

        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar2.Value = e.ProgressPercentage;
            label5.Text = "Messages sent: " + e.ProgressPercentage + "%";
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MessageBox.Show("Successfully cleared message-queue");
            }
            else
            {
                stopwatch.Stop();
                TimeSpan ts = stopwatch.Elapsed;
                string elapsedTime = string.Format("{0:00}:{1:00}:{2:00}.{3:00}", ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
                int total = amount * checkedListBox1.CheckedItems.Count;
                MessageBox.Show("Messages were successfully sent" + Environment.NewLine + "Total amount of messages sent: " + total + Environment.NewLine + "Elapsed Time: " + elapsedTime + Environment.NewLine + " Average Messages per Second: " + Math.Round((total / ts.TotalMilliseconds) * 1000));
                stopwatch.Reset();
                amount = 0;

            }

        }
        private void sendMessages()
        {
            for (int a = 0; a <= amount; a++)
            {
                int i = 0;
                while (i < checkedListBox1.CheckedItems.Count)
                {
                    skype.SendMessage(checkedListBox1.CheckedItems[i].ToString(), message);
                    i++;
                }

                int progress = a * 100 / amount;
                backgroundWorker1.ReportProgress(progress);
                if (backgroundWorker1.CancellationPending)
                {
                    backgroundWorker1.ReportProgress(100);
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            backgroundWorker1.CancelAsync();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            PictureBox1.Image = new System.Drawing.Bitmap(new System.IO.MemoryStream(new System.Net.WebClient().DownloadData("http://api.skype.com/users/" + textBox2.Text + "/profile/avatar")));
            User cuser = skype.User[textBox2.Text];
            textBox3.Text = cuser.FullName;
            textBox4.Text = cuser.RichMoodText;
            textBox5.Text = cuser.Country;
            textBox6.Text = cuser.City;
            textBox7.Text = cuser.Birthday.ToString();
            textBox8.Text = cuser.LastOnline.ToString();
            if (cuser.HasCallEquipment)
            {
                textBox9.Text = "Yes";
            }
            else
            {
                textBox9.Text = "No";
            }
            if(cuser.IsVideoCapable)
            {
                textBox10.Text = "Yes";
            }
            else
            {
                textBox10.Text = "No";
            }
            textBox11.Text = cuser.About;
            textBox12.Text = cuser.Aliases;
            textBox13.Text = cuser.CountryCode;
            textBox14.Text = cuser.Language;
            textBox15.Text = cuser.Timezone.ToString();
            textBox16.Text = cuser.Sex.ToString();
            String path = "C:/Users/cedga/Desktop";

        }

        private void button12_Click(object sender, EventArgs e)
        {
            DataObject data = QuoteUtils.getObject(textBox18.Text, richTextBox2.Text, textBox17.Text);
            Clipboard.SetDataObject(data, true);

        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            textBox17.Text = DateTime.Now.ToString();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            switch (checkBox1.Checked)
            {
                case true:
                    timer3.Enabled = true;
                    break;
                case false:
                    timer3.Enabled = false;
                    break;
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            switch (radioButton1.Checked)
            {
                case true:
                    textBox17.Enabled = false;
                    break;
                case false:
                    textBox17.Enabled = true;
                    break;
            }

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            switch (radioButton2.Checked)
            {
                case true:
                    textBox17.Enabled = true;
                    break;
                case false:
                    textBox17.Enabled = false;
                    break;
            }

        }
    }
}

