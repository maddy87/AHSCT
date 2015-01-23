using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;

namespace maddytry1
{
    public partial class Feedback : Form
    {
        public string sSubject;
        public Feedback()
        {
            InitializeComponent();
        }

        private void rdbFeedback_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbFeedback.Checked == true)
            {
                sSubject = "AHSCT: FEEDBACK";
            }

        }

        private void rdbBug_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbBug.Checked == true)
            {
                sSubject = "AHSCT: BUG";
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            #region: EMAILING THE STATUS OF THE CURRENT USERS
            if (txtMessage.Text == "")
            {
                MessageBox.Show(" Please provide and appropriate feedback ", "No Empty Feebback Plz", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                try
                {
                    string sysname = System.Environment.MachineName.ToString();
                    string uid1 = System.Environment.UserName.ToString();

                    //MessageBox.Show("System NAme : " +sysname+ " UID : " +uid + " UserName : " +uid1+ "  " + GlobalData.gCurrentUser);

                    MailMessage Send_Info = new MailMessage();
                    Send_Info.From = new MailAddress(uid1 + "@Ikran.com");
                    Send_Info.To.Add("rajesh.shetty@Ikran.com");
                    Send_Info.Subject = sSubject;
                    Send_Info.Body = txtMessage.Text;
                    SmtpClient client = new SmtpClient("172.19.98.22", 25);
                    client.UseDefaultCredentials = true;

                    //client.Send(Send_Info);
                    this.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error Sending Email", "Problem Sending Mail", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            #endregion
        }
    }
}
