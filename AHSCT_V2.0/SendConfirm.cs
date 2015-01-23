using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net.Mail;
//using Outlook = Microsoft.Office.Interop.Outlook;

namespace maddytry1
{
    public partial class SendConfirm : Form
    {
        public SendConfirm()
        {
            InitializeComponent();
        }

        private void SendConfirm_Load(object sender, EventArgs e)
        {
            lblStatus.Visible = false;
             //txtFrom.Text = "rajesh_shetty02@Ikran.com";
             //txtTo.Text = "rajesh_shetty02@Ikran.com";
             txtTo.Text = GlobalData.gSeverity == 1 ? GlobalData.gSeverity1DL : GlobalData.gSeverity2DL;
             txtCc.Text = "rajesh_shetty02@Ikran.com";
             if (GlobalData.gNoti_Type == "N" || GlobalData.gNoti_Type == "F" || GlobalData.gNoti_Type == "FF" || GlobalData.gNoti_Type == "FN")
             {
                 txtSubject.Text = GlobalData.gStatus + " : Severity " + GlobalData.gSeverity + " Production Incident : " + GlobalData.gApplicatioName + " Ticket Reference : " + GlobalData.gCurrTicketRef;
             }
             else
             {
                 txtSubject.Text = GlobalData.gStatus + " " + GlobalData.gUpNo + " : Severity " + GlobalData.gSeverity + " Production Incident : " + GlobalData.gApplicatioName + " Ticket Reference : " + GlobalData.gCurrTicketRef;
             }
             //MessageBox.Show(GlobalData.gType);
        }

        private void btSend_Click(object sender, EventArgs e)
        {

            #region: INSERTING DATA INTO MASTER,CONSOLE AND AVAILABILITY TABLES

            lblStatus.Text = "Updating Details into Database  : 27%";

            //GlobalData.gType = "Final";
            Functions obj = new Functions();
            
            // INSERTING THE DATA INTO THE MASTER TABLE (RECORD FOR EVERY NOTIFICATION CREATED)
            obj.InsertDetails(GlobalData.gCurrTicketRef, GlobalData.gApplicatioName, GlobalData.gType, GlobalData.gUpNo, GlobalData.gSeverity,
                              GlobalData.gIncTitle, GlobalData.gBriefSumm, GlobalData.gSummPart1, GlobalData.gSummPart2, GlobalData.gDetails,
                              GlobalData.gBIActual, GlobalData.gBIPotential, GlobalData.gPreviousActions, GlobalData.gStartDT, GlobalData.gNextDT, GlobalData.gPOCEmailId);
            
            //MessageBox.Show("MASTER DATABASE UPDATED");
            lblStatus.Text = "Updating Details into Database : 53%";

            //IF THE RECORD IS BEING CREATED FOR THE FIRST TIME IN THE DATABASE.
            if (GlobalData.gNoti_Type == "N" || GlobalData.gNoti_Type == "FN")
            {
                obj.InsertCurrentDetails(GlobalData.gCurrTicketRef, GlobalData.gApplicatioName, GlobalData.gType, GlobalData.gUpNo,
                                         GlobalData.gSeverity, GlobalData.gIncTitle, GlobalData.gBriefSumm, GlobalData.gSummPart1,
                                         GlobalData.gSummPart2, GlobalData.gDetails, GlobalData.gBIActual, GlobalData.gBIPotential,
                                         GlobalData.gPreviousActions, GlobalData.gStartDT, GlobalData.gNextDT, GlobalData.gPOCEmailId);
                //MessageBox.Show("RECORDS INSERTED INTO CURRENT DATABASE");
                lblStatus.Text = "Updating Details into Database : 97%";
            }

            //IF THE RECORD IS ALREADY PRESENT THEN UPDATE THAT DATA IN THE DATABASE.
            else
            {
                obj.UpdateCurrentDetails(GlobalData.gCurrTicketRef, GlobalData.gType, GlobalData.gUpNo, GlobalData.gSeverity, GlobalData.gNextDT);
                //MessageBox.Show("RECORDS UPDATED INTO CURRENT DATABASE");
                lblStatus.Text = "Updating Details into Database : 99%";
            }

            #endregion

            #region UPDATING THE DETAILS IN AVAILABILITY TRACKER

            if (GlobalData.gUpdateAT == 1)
            {
                Functions fobj = new Functions();
                fobj.UpdateAvailabilityTracker(GlobalData.gStartDT, GlobalData.gTempNextDT, GlobalData.gOutageDuration, GlobalData.gPlannedActivity, GlobalData.gDueToAM, GlobalData.gActionOwner);
                fobj = null;
                //MessageBox.Show("DETAILS UPDATED IN THE AVAILABILITY TRACKER");
                lblStatus.Text = "Updating Details into Database : 100%";
            }
            #endregion

            #region :SENDING MAIL USING MACRO
            //SENDING THE MAIL TO THE DESIRED DL
            try
            {
                string mailto = txtTo.Text;
                string mailfrom = txtFrom.Text;
                string mailcc = txtCc.Text;
                string subject = txtSubject.Text;

                Microsoft.Office.Interop.Word.Application word1 = null;
                Microsoft.Office.Interop.Word.Document doc1 = null;
                object filename = GlobalData.gFilename;
                object oFrom = mailfrom;
                object oTo = mailto;
                object oCc = mailcc;
                object oSubject = subject;
                object missing1 = System.Reflection.Missing.Value;
                object readtrue = true;
                word1 = new Microsoft.Office.Interop.Word.Application();
                doc1 = new Microsoft.Office.Interop.Word.Document();

                word1.Visible = false;
                doc1 = word1.Documents.Open(ref filename, ref missing1, ref readtrue, ref missing1, ref missing1, ref missing1, ref missing1, ref missing1,
                                           ref missing1, ref missing1, ref missing1, ref missing1, ref missing1, ref missing1, ref missing1, ref missing1);



                //Running the appropirate macro from the word file

                word1.Application.Run("TestMail", ref oFrom, ref oTo, ref oCc, ref oSubject, ref missing1, ref missing1, ref missing1, ref missing1, ref missing1, ref
                                   missing1, ref missing1, ref missing1, ref missing1, ref missing1, ref missing1, ref missing1, ref missing1, ref missing1, ref missing1, ref
                                   missing1, ref missing1, ref missing1, ref missing1, ref missing1, ref missing1, ref missing1, ref missing1, ref missing1, ref missing1, ref missing1);

                //RunMacro(word1, new object[] { "TestMail", mailfrom, mailto, mailcc, subject });

                //Closing the word app

                object saveChanges = false;
                object originalFormat = System.Reflection.Missing.Value;
                object routeDocument = System.Reflection.Missing.Value;

                word1.Quit(ref saveChanges, ref originalFormat, ref routeDocument);
            }

            catch (Exception ex)
            {
                Functions fobj = new Functions();
                fobj.ErrorReporting(ex.Message.ToString(), "#region :SENDING MAIL USING MACRO , btn_Send", "127");
            }
            finally
            {
                //REPORTING THE INFORMATION TO THE ME .

                try
                {
                    string sysname = System.Environment.MachineName.ToString();
                    string uid1 = System.Environment.UserName.ToString();

                    //MessageBox.Show("System NAme : " +sysname+ " UID : " +uid + " UserName : " +uid1+ "  " + GlobalData.gCurrentUser);

                    MailMessage Send_Info = new MailMessage();
                    Send_Info.From = new MailAddress(uid1 + "@Ikran.com");
                    Send_Info.To.Add("rajesh.shetty@Ikran.com");
                    Send_Info.Subject = "AHSCT:Notify - I have sent the " + GlobalData.gNoti_Type + " - " +GlobalData.gCurrTicketRef+ " - " + GlobalData.gApplicatioName;
                    Send_Info.Body = "System Name : " + sysname + Environment.NewLine + "User Name : " + uid1;
                    SmtpClient client = new SmtpClient("172.19.98.22", 25);
                    client.UseDefaultCredentials = true;
                    //client.Send(Send_Info);

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error Sending Email", "Problem Sending Mail", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                }
            }
            #endregion

            #region :SENDING MAIL USING OUTLOOK (ELIMINATING THE NEED OF MACRO) 

            //Microsoft.Office.Interop.Outlook.ApplicationClass olkApp = new Microsoft.Office.Interop.Outlook.ApplicationClass();
            ////olkApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
            //Outlook.MailItem olkMailItem = (Outlook.MailItem)olkApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
            //olkMailItem.To = "rajesh_shetty02@Ikran.com";
            //olkMailItem.Subject = "AS12_Test_Notification";
            //olkMailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
            //olkMailItem.HTMLBody = GlobalData.gEmailBody;
            //olkMailItem.Send();

            #endregion
            
            #region: RELOADING ALL THE DATAGRIDS

            

            #endregion

            //UPDATING THE CURRENT INFO 
            GlobalData.gUpdateCurrentInfo = 1;
            GlobalData.gHasNotificationSent = 1;
            this.Dispose();
            this.Close();

           
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            //this.Dispose();
            this.Close();

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}
