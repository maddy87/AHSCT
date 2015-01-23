using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Threading;

namespace maddytry1
{
    public partial class Generate : Form
    {

        //Date Variables
        DateTime CurrentDT = new DateTime();
        DateTime StartTimeDT = new DateTime();
        DateTime NextUpdateDT = new DateTime();
        DateTime EstimatedResolutionDT = new DateTime();
        Operations obj = new Operations();

        //TIMING VARIABLES
        Double Offsettminutes = 30;
        Double OffsettResolutionMins = 90;
        Double Offsetthours = 2;

        public Generate()
        {
            InitializeComponent();
        }

        private void Generate_Load(object sender, EventArgs e)
        {
            #region: LOADING THE APPLICATION RELATED INFORMATION INTO THE GLOBAL VARIABLES
            Thread th_GetAppInfo = new Thread(new ThreadStart(obj.GetApplicationInfo));
            th_GetAppInfo.Start();
            #endregion

            #region DEFAULTING VALUES

            //DEFAULTING DATE TIME VALUES
            CurrentDT = DateTime.Now;

            //DEFAULTING AVAILABILITY TRACKER VALUES
            GlobalData.gPlannedActivity = "N";
            GlobalData.gDueToAM = "N";
            chkAT.Visible = false;

            //Prepare the form according to the data recieved.

            grpSevLowered.Visible = false;
            grpAvailTracker.Visible = false;

            //PREPARE THE HEADER TEXT ACCORDING TO THE DATA RECEIVED.
            if(GlobalData.gNoti_Type == "N")
            {
             lbHeader.Text =  GlobalData.gStatus + " - " + GlobalData.gApplicatioName;
            }

            else
            {
             lbHeader.Text =  GlobalData.gStatus + " " + GlobalData.gUpNo + " - " + GlobalData.gApplicatioName;
            }
            rdbSev1.Checked = GlobalData.gSeverity == 1 ? true : false;
            rdbSev2.Checked = GlobalData.gSeverity == 2 ? true : false;


            txtSend.Visible = false;
            lblRDBErr.Visible = false;

            //PUPULATING THE TIMEZONE LABELS

            lblSTZ.Text = GlobalData.gTimeZone;
            lblSTZ.Visible = false;
            cmbTimZone.SelectedText = GlobalData.gTimeZone;
            cmbTimZone.Enabled = false;
            lblNTZ.Text = GlobalData.gTimeZone;
            lblETZ.Text = GlobalData.gTimeZone;
            chkLocalTZ.Visible = false;
            lblLocalTZ.Visible = false;
            txtLTZ.Visible = false;
            btnSetTZ.Visible = false;
            lblexTZ.Visible = false;


            #endregion

            #region IN CASE IT IS A NOTIFICATION
            //IF IT A NEW NOTIFICATION BEING CREATED.
            if (GlobalData.gNoti_Type == "N" || GlobalData.gNoti_Type == "FN") 
            {
                //CONTROLS NOT TO BE VISIBLE WHEN  CREATING A NOTIFICATION.
                //lblSummp2.Visible = false;
                //txtSummP2.Visible = false; // I NEED THESE AS IF NOT USED IT WILL LEAVE THE bkSummaryPart3 BOOKMARK EMPTY
                lblSummp1.Text = "Latest Update";

                StartTimeDT = CurrentDT; // As this is the firstnotification.
                NextUpdateDT = CurrentDT.AddMinutes(Offsettminutes);
                EstimatedResolutionDT = CurrentDT.AddHours(Offsetthours);

                txtStartDT.Text = StartTimeDT.ToString("dd MMMM yyyy HH:mm"); //+ " " + GlobalData.gTimeZone;
                GlobalData.gTempStartDT = StartTimeDT.ToString();

                txtNextDT.Text = NextUpdateDT.ToString("dd MMMM yyyy HH:mm"); //+ " " + GlobalData.gTimeZone;
                GlobalData.gTempNextDT = NextUpdateDT.ToString();

                txtEstimatedResolution.Text = EstimatedResolutionDT.ToString("dd MMMM yyyy HH:mm"); //+ " " + GlobalData.gTimeZone;
                GlobalData.gTempResolvedDT = EstimatedResolutionDT.ToString();

                //VALIDATIONS PENDING.
                
                if (GlobalData.gNoti_Type == "FN")
                {
                    string gstatus = GlobalData.gStatus;
                    GlobalData.gUpNo = 0;
                    txtIncidentTitle.Text = GlobalData.gIncTitle;
                    txTicketRef.Text = GlobalData.gCurrTicketRef;
                    txtBriefSummary.Text = GlobalData.gBriefSumm;
                    txtSummP1.Text = GlobalData.gSummPart1;
                    txtSummP2.Text = " ";
                    txtDetails.Text = GlobalData.gDetails;
                    txtBIP.Text = GlobalData.gBIPotential;
                    txtBIA.Text = GlobalData.gBIActual;
                    txtPreviousActions.Text = GlobalData.gPreviousActions;
                    //txtStartDT.Text = GlobalData.gStartDT; ... 
                    //txtNextDT.Text = GlobalData.gNextDT;
                    txtPOCName.Text = GlobalData.gPOCName;
                    txtPOCContact.Text = "9158889692";
                    txtPOCEmail.Text = "rajesh.shetty@Ikran.com";

                    GlobalData.gNoti_Type = "N";
                }

            }
            #endregion

            #region IF IT IS AND UPDATE/FINAL NOTIFICATION BEING CREATED
            //IF IT IS AND UPDATE/FINAL NOTIFICATION BEING CREATED
            else
            {
                //IF IT IS AND UPDATE 
                if (GlobalData.gNoti_Type == "U")
                {
                    //UPDATING THE FORM WITH THE WITHDRAWN VALUES
                    txtIncidentTitle.Text = GlobalData.gIncTitle;
                    txTicketRef.Text = GlobalData.gCurrTicketRef;
                    txtBriefSummary.Text = GlobalData.gBriefSumm;
                    txtSummP1.Text = GlobalData.gSummPart1;
                    txtSummP2.Text = "---Next Planned Actions Details Here--";//GlobalData.gSummPart2;
                    txtDetails.Text = GlobalData.gDetails;
                    txtBIP.Text = GlobalData.gBIPotential;
                    txtBIA.Text = GlobalData.gBIActual;
                    txtPreviousActions.Text = GlobalData.gPreviousActions;
                    
                    //CONVERTING THE ACCESS FORMAT DATE INTO THE NOTIFCATION FORMAT DATE
                    CultureInfo culture = new CultureInfo("en-US");
                    DateTime dtTemp = Convert.ToDateTime(GlobalData.gStartDT, culture);
                    txtStartDT.Text = dtTemp.ToString("dd MMMM yyyy HH:mm"); // +" " + GlobalData.gTimeZone; 
                    GlobalData.gTempStartDT =  GlobalData.gStartDT;

                    //dtTemp = Convert.ToDateTime(GlobalData.gNextDT, culture);
                    dtTemp = DateTime.Now.AddMinutes(Offsettminutes);
                    txtNextDT.Text = dtTemp.ToString("dd MMMM yyyy HH:mm"); // +" " + GlobalData.gTimeZone;
                    GlobalData.gTempNextDT = dtTemp.ToString();

                    dtTemp = DateTime.Now.AddHours(Offsetthours);
                    txtEstimatedResolution.Text = dtTemp.ToString("dd MMMM yyyy HH:mm"); // +" " + GlobalData.gTimeZone;
                    GlobalData.gTempResolvedDT = dtTemp.ToString();

                    txtPOCName.Text = GlobalData.gPOCName;
                    txtPOCContact.Text = "9158889692";
                    txtPOCEmail.Text = "rajesh.shetty@Ikran.com";

                    //DateTime temp = new DateTime();
                    //DateTime temp = DateTime.Now;
                    //temp = Convert.ToDateTime(GlobalData.gNextDT);
                    //txtNextDT.Text = temp.AddMinutes(Offsettminutes).ToString();
                    //txtEstimatedResolution.Text = temp.AddHours(Offsetthours).ToString();
                }
            #endregion

            #region IF IT IS AN FINAL NOTIFICATION
                else
                {
                    if (GlobalData.gNoti_Type == "FF")
                    {
                        chkAT.Visible = true;
                        //lblNextDT.Visible = false;
                        //txtNextDT.Visible = false;
                        txtStartDT.Text = CurrentDT.ToString("dd MMMM yyyy HH:mm"); // +" " + GlobalData.gTimeZone;
                        GlobalData.gTempStartDT = CurrentDT.ToString();
                        //txtEstimatedResolution.Text = CurrentDT.AddMinutes(Offsettminutes).ToString();
                        //lblEstimatedResolution.Text = "Resolution Date/Time";
                        lblBIP.Visible = false;
                        txtBIP.Visible = false;
                        lblBIA.Text = "Impact - Actual";
                        lblNextDT.Text = "Resolution Date/Time";
                        //txtNextDT.Text = CurrentDT.AddMinutes(Offsettminutes).ToString();
                        //RESOLUTION TIME SET TO CURRENT TIME BY DEFAULT
                        txtNextDT.Text = DateTime.Now.ToString("dd MMMM yyyy HH:mm"); // +" " + GlobalData.gTimeZone;
                        GlobalData.gTempNextDT = DateTime.Now.ToString();

                        lblEstimatedResolution.Visible = false;
                        txtEstimatedResolution.Visible = false;
                        lblETZ.Visible = false;
                           
                     
                    }
                    else
                    {
                        //IF SENDING THE FINAL UPDATE

                        lblBIP.Visible = false;
                        txtBIP.Visible = false;
                        lblBIA.Text = "Impact - Actual";

                        lblNextDT.Text = "Resolution Date/Time";
                        lblEstimatedResolution.Visible = false;
                        txtEstimatedResolution.Visible = false;
                        lblETZ.Visible = false;

                        grpSevLowered.Visible = true;
                        //grpAvailTracker.Visible = true;  //WILL BE DISPLAYED WHEN THE CHECKBOX IS CHECKED.
                        //FILL DETAILS IN THE AVAILABILITY TRACKER
                        chkAT.Visible = true;

                        rdbSev3.Visible = false;
                        rdbSev4.Visible = false;

                        txtIncidentTitle.Text = GlobalData.gIncTitle;
                        txTicketRef.Text = GlobalData.gCurrTicketRef;
                        txtBriefSummary.Text = GlobalData.gBriefSumm;
                        txtSummP1.Text = GlobalData.gSummPart1;
                        txtSummP2.Text = GlobalData.gSummPart1;
                        txtDetails.Text = GlobalData.gDetails;
                        //txtBIP.Text = GlobalData.gBIPotential;
                        txtBIA.Text = GlobalData.gBIActual;
                        txtPreviousActions.Text = GlobalData.gPreviousActions;

                        CultureInfo culture = new CultureInfo("en-US");
                        DateTime dtTemp = Convert.ToDateTime(GlobalData.gStartDT, culture);
                        txtStartDT.Text = dtTemp.ToString("dd MMMM yyyy HH:mm"); // +" " + GlobalData.gTimeZone;
                        GlobalData.gTempStartDT = GlobalData.gStartDT;
                        
                        //txtNextDT.Text = GlobalData.gNextDT;
                        txtPOCName.Text = GlobalData.gPOCName;
                        txtPOCContact.Text = "9158889692";
                        txtPOCEmail.Text = "rajesh.shetty@Ikran.com";

                        //DateTime temp = new DateTime();
                        //temp = Convert.ToDateTime(GlobalData.gNextDT);
                        //txtNextDT.Text = temp.AddMinutes(Offsettminutes).ToString();
                        //RESOLUTION TIME SET TO CURRENT TIME BY DEFAULT
                        txtNextDT.Text = DateTime.Now.ToString("dd MMMM yyyy HH:mm"); // +" " + GlobalData.gTimeZone;
                        GlobalData.gTempNextDT = DateTime.Now.ToString();

                        //POPULATING THE AVAILABILITY TRACKER DETAILS ACCORDINGLY
                        txtUAStartDate.Text = dtTemp.ToString("dd MMMM yyyy HH:mm"); // +" " + GlobalData.gTimeZone;
                        txtUAEndDate.Text = DateTime.Now.ToString("dd MMMM yyyy HH:mm"); // +" " + GlobalData.gTimeZone;
                        txtActionOwner.Text = "Ikran";

                    }
                }
                #endregion

            //IRRESPECTIVE IF IT IS AN UPDATE OR AN FINAL ONE.
            }

            #region: LOADING INFO INTO THE STATUS BAR

            stslblClass.Text = GlobalData.gServiceClass + " " + "Application" ;
            stslblCountry.Text = GlobalData.gCountryOfSupport;
            stslblUserBaseData.Text = " " + GlobalData.gUserBase + " ";
            stslblClientArea.Text = "Client Area : " + GlobalData.gClientArea + " ";
            stslblONL2M.Text = " " + GlobalData.ONL2Manager + " ";
            stsOFL2M.Text = " " + GlobalData.OFFL2Manager + " ";
            
            #endregion

            //DEFAULTING VALUES 
            txtDetails.Text = "Incident Started at " + txtStartDT.Text + " " + lblSTZ.Text;
        }

        private void rdbSev1_CheckedChanged(object sender, EventArgs e)
        {
            GlobalData.gSeverity = rdbSev1.Checked == true ?  1 :  2;
        }

        private void rdbSev2_CheckedChanged(object sender, EventArgs e)
        {
            GlobalData.gSeverity = rdbSev2.Checked == true ? 2 : 1;
        }

        private void chkLowerSeverity_CheckedChanged(object sender, EventArgs e)
        {
            if (chkLowerSeverity.Checked == true)
            {
                rdbSev3.Visible = true;
                rdbSev4.Visible = true;
            }

            else
            {
                rdbSev3.Visible = false;
                rdbSev4.Visible = false;
            }
        }

       
        public void SpellCheck()
        {

            Microsoft.Office.Interop.Word.Application spellword = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document spelldoc = new Microsoft.Office.Interop.Word.Document();

            try
            {
                
                spellword.Visible = false;
                object filename = GlobalData.gFilename;
                object falsevalue = false;
                object truevalue = true;
                object missing = System.Reflection.Missing.Value;
                object template = System.Reflection.Missing.Value;

                spelldoc = spellword.Documents.Add(ref template, ref template, ref missing, ref truevalue);


                //Incident Title

                spelldoc.Words.First.InsertBefore(txtIncidentTitle.Text);
                Word.ProofreadingErrors we = spelldoc.SpellingErrors;
                int ErrorCount = we.Count;
                spelldoc.CheckSpelling(ref missing, ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                object first = 0;
                object last = spelldoc.Characters.Count - 1;
                txtIncidentTitle.Text = spelldoc.Range(ref first, ref last).Text;
                spellword.Visible = false;
                spelldoc.Range(ref first, ref last).Cut();

                //Ticket Reference

                spelldoc.Words.First.InsertBefore(txTicketRef.Text);
                we = spelldoc.SpellingErrors;
                ErrorCount = we.Count;
                spelldoc.CheckSpelling(ref missing, ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                first = 0;
                last = spelldoc.Characters.Count - 1;
                txTicketRef.Text = spelldoc.Range(ref first, ref last).Text;
                spellword.Visible = false;
                spelldoc.Range(ref first, ref last).Cut();

                ////Brief summary

                spelldoc.Words.First.InsertBefore(txtBriefSummary.Text);
                we = spelldoc.SpellingErrors;
                ErrorCount = we.Count;
                spelldoc.CheckSpelling(ref missing, ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                first = 0;
                last = spelldoc.Characters.Count - 1;
                txtBriefSummary.Text = spelldoc.Range(ref first, ref last).Text;
                spellword.Visible = false;
                spelldoc.Range(ref first, ref last).Cut();

                ////Summary Part 1
                spelldoc.Words.First.InsertBefore(txtSummP1.Text);
                we = spelldoc.SpellingErrors;
                ErrorCount = we.Count;
                spelldoc.CheckSpelling(ref missing, ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                first = 0;
                last = spelldoc.Characters.Count - 1;
                txtSummP1.Text = spelldoc.Range(ref first, ref last).Text;
                spellword.Visible = false;
                spelldoc.Range(ref first, ref last).Cut();


                //Summary Part 2
                if (GlobalData.gNoti_Type != "FN")
                {
                    spelldoc.Words.First.InsertBefore(txtSummP2.Text);
                    we = spelldoc.SpellingErrors;
                    ErrorCount = we.Count;
                    spelldoc.CheckSpelling(ref missing, ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                    first = 0;
                    last = spelldoc.Characters.Count - 1;
                    txtSummP2.Text = spelldoc.Range(ref first, ref last).Text;
                    spellword.Visible = false;
                    spelldoc.Range(ref first, ref last).Cut();
                }
                ////Details

                spelldoc.Words.First.InsertBefore(txtDetails.Text);
                we = spelldoc.SpellingErrors;
                ErrorCount = we.Count;
                spelldoc.CheckSpelling(ref missing, ref missing, ref missing, ref missing,
                                      ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                first = 0;
                last = spelldoc.Characters.Count - 1;
                txtDetails.Text = spelldoc.Range(ref first, ref last).Text;
                spellword.Visible = false;
                spelldoc.Range(ref first, ref last).Cut();

                ////BI - ACTUAL

                spelldoc.Words.First.InsertBefore(txtBIA.Text);
                we = spelldoc.SpellingErrors;
                ErrorCount = we.Count;
                spelldoc.CheckSpelling(ref missing, ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                first = 0;
                last = spelldoc.Characters.Count - 1;
                txtBIA.Text = spelldoc.Range(ref first, ref last).Text;
                spellword.Visible = false;
                spelldoc.Range(ref first, ref last).Cut();

                if (GlobalData.gNoti_Type == "N" || GlobalData.gNoti_Type == "U" || GlobalData.gNoti_Type == "FN")
                {
                    ////BI - POTENTIAL

                    spelldoc.Words.First.InsertBefore(txtBIP.Text);
                    we = spelldoc.SpellingErrors;
                    ErrorCount = we.Count;
                    spelldoc.CheckSpelling(ref missing, ref missing, ref missing, ref missing,
                                            ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                    first = 0;
                    last = spelldoc.Characters.Count - 1;
                    txtBIP.Text = spelldoc.Range(ref first, ref last).Text;
                    spellword.Visible = false;
                    spelldoc.Range(ref first, ref last).Cut();
                }
                ////PREVIOUS ACTIONS

                spelldoc.Words.First.InsertBefore(txtPreviousActions.Text);
                we = spelldoc.SpellingErrors;
                ErrorCount = we.Count;
                spelldoc.CheckSpelling(ref missing, ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                first = 0;
                last = spelldoc.Characters.Count - 1;
                txtPreviousActions.Text = spelldoc.Range(ref first, ref last).Text;
                spellword.Visible = false;
                spelldoc.Range(ref first, ref last).Cut();

                //INCLUDING SPELLCHECK FOR THE RECENT DATE FORMATT

                spelldoc.Words.First.InsertBefore(txtStartDT.Text);
                we = spelldoc.SpellingErrors;
                ErrorCount = ErrorCount + we.Count;
                spelldoc.CheckSpelling(ref missing, ref missing, ref missing, ref missing,
                                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                first = 0;
                last = spelldoc.Characters.Count - 1;
                txtStartDT.Text = spelldoc.Range(ref first,ref last).Text;
                spelldoc.Range(ref first, ref last).Cut();

                //QUITING OUT OF THE WORD FILE 
                object savechanges = false;
                spellword.Quit(ref savechanges, ref missing, ref missing);
                spelldoc = null;
                spellword = null;
                GC.Collect();

                //DISPLAYING THE NO OF ERRORS
                if (ErrorCount == 0)
                    lblErrors.Text = "Spelling OK. No errors corrected ";
                else if (ErrorCount == 1)
                    lblErrors.Text = "Spelling OK. 1 error corrected ";
                else
                    lblErrors.Text = "Spelling OK. " + ErrorCount + " errors corrected ";

                //TRIMMING THE EXTRA SPACES

                txtIncidentTitle.Text = txtIncidentTitle.Text.Trim();
                txtBriefSummary.Text = txtBriefSummary.Text.Trim();
                txtSummP1.Text = txtSummP1.Text.Trim();
                txtSummP2.Text = txtSummP2.Text.Trim();
                txtDetails.Text = txtDetails.Text.Trim();
                txtBIA.Text = txtBIA.Text.Trim();
                txtBIP.Text = txtBIP.Text.Trim();
                txtPreviousActions.Text = txtPreviousActions.Text.Trim();
                btnPreviewNotification.Text = "GENERATING NOTIFICATION";
            }
            catch (Exception exx)
            {
                MessageBox.Show("Error Performing SpellCheck", "Misspelled Something");
            }
            finally
            {
              //  Marshal.ReleaseComObject(spelldoc);
                
            }

        }


        
        private void btnPreviewNotification_Click(object sender, EventArgs e)
        {
            #region: VALIDATIONS OF THE PROVIDED INPUTS
            try
            {
                ValidateData(); 
                ValidateDate();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem Validating Data. PLease check the information you have provided and try again", "In-Valid Data");
            }
                if (GlobalData.gValidate == 1 || GlobalData.gValidateDate == 1)
                {
                    //DO NOTHING AS INVALID DATA HAS BEEN PROVIDED.
                    GlobalData.gValidateDate = 0;
                }
            

             #endregion
                             
            else
            {
                #region: SET THE DATE'S ACCORDING TO THE ONE'S ENTERED IN THE TEXT BOXES

                CultureInfo cul = new CultureInfo("en-US");
                DateTime dtStart = Convert.ToDateTime(txtStartDT.Text,cul);
                DateTime dtNext = Convert.ToDateTime(txtNextDT.Text, cul);
                GlobalData.gActionOwner = txtActionOwner.Text;
                GlobalData.gTempStartDT = dtStart.ToString();
                GlobalData.gTempNextDT = dtNext.ToString();
                GlobalData.gStartDT = GlobalData.gTempStartDT;
                GlobalData.gNextDT = GlobalData.gTempNextDT;

                #endregion

                txtSend.Visible = true;


                if (GlobalData.gNoti_Type == "N" || GlobalData.gNoti_Type == "FN")
                {
                    GlobalData.gCurrTicketRef = txTicketRef.Text;
                    GlobalData.gIncTitle = txtIncidentTitle.Text;
                    GlobalData.gCurrTicketRef = txTicketRef.Text;
                    GlobalData.gBriefSumm = txtBriefSummary.Text;
                    GlobalData.gSummPart1 = txtSummP1.Text;
                    GlobalData.gSummPart2 = txtSummP2.Text;
                    GlobalData.gDetails = txtDetails.Text;
                    GlobalData.gBIPotential = txtBIP.Text;
                    GlobalData.gBIActual = txtBIA.Text;
                    GlobalData.gPreviousActions = txtPreviousActions.Text;
                    GlobalData.gStartDT = GlobalData.gTempStartDT;
                    GlobalData.gNextDT = GlobalData.gTempNextDT;
                    GlobalData.gPOCName = txtPOCName.Text;
                    GlobalData.gPOCEmailId = txtPOCEmail.Text;
                
                    //txtPOCEmail.Text = "rajesh_shetty02@Ikran.com";
                }

                //Creating the location for the notifications generated 
                
                //System.IO.Directory.CreateDirectory(GlobalData.gFileLoaction);

                //AZ File Location
                GlobalData.gFileLoaction = System.Windows.Forms.Application.StartupPath + "\\" + GlobalData.gCategory + "\\" + GlobalData.gApplicatioName + "\\" + GlobalData.gCurrTicketRef + "";
                System.IO.Directory.CreateDirectory(GlobalData.gFileLoaction);


                //Preparing the filename
                GlobalData.gFilename = GlobalData.gFileLoaction + "\\" + GlobalData.gStatus + "-" + GlobalData.gUpNo + "-" + GlobalData.gCurrTicketRef + "-" + GlobalData.gApplicatioName + ".doc";
                if (GlobalData.gNoti_Type == "FF")
                {
                    GlobalData.gFilename = GlobalData.gFileLoaction + "\\Notification_Final-" + GlobalData.gCurrTicketRef + "-" + GlobalData.gUpNo + "-" + GlobalData.gApplicatioName + ".doc";
                }

                // PERFORMING THE SPELLCHECK AS DESIRED.
                btnPreviewNotification.Text = "PERFORMING SPELLCHECK";
                SpellCheck();

                #region : GENERATING THE NOTIFICATION

                Microsoft.Office.Interop.Word.Application genword = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word.Document gendoc = new Microsoft.Office.Interop.Word.Document();


                genword.Visible = true;
                object template_filename;
                object filename = GlobalData.gFilename;
                object falsevalue = false;
                object truevalue = true;
                object missing = System.Reflection.Missing.Value;

                //MessageBox.Show("Entering Try");
                try
                {

                    //SELECTING THE TEMPLATE HERE

                    if (GlobalData.gNoti_Type == "F" || GlobalData.gNoti_Type == "FF")
                    {
                        //template_filename = "D:\\Sample_Final.docm";
                        
                        //ONE MORE TIME FOR AZ
                        template_filename = System.Windows.Forms.Application.StartupPath.ToString() + "\\Sample_F.doc";
                        //MessageBox.Show("Initializing Templates");
                    }

                    else
                    {
                        //template_filename = "D:\\Sample_NotiUpdate.doc";
                        
                        template_filename = System.Windows.Forms.Application.StartupPath.ToString() + "\\Sample_NU.doc";
                        //MessageBox.Show("Initializing Templates");
                    }


                    gendoc = genword.Documents.Open(ref template_filename, ref missing, ref truevalue, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);



                    //MessageBox.Show("Opening document for start creating template");
                    //CREATING BOOKMARK OBJECT COMMON TO BOTH THE TEMPLATES

                    object obkSev = "bkSeverity";
                    object obkIncTil = "bkIncidentTitle";
                    object obkBrSum = "bkBriefSummary";
                    object obkSummP2 = "bkSummaryPart2";
                    object obkSummP3 = "bkSummaryPart3";
                    object obkTckRef = "bkTicketReference";
                    object obkListApp = "bkListAppAffected";
                    object obkDetails = "bkDetails";
                    object obkPrevAct = "bkPreviousActions";
                    object obkPOC = "bkPointOfContact";

                    //MessageBox.Show("Initializing Bookmarks");  
                    //INTEROP RANGE 
                    Microsoft.Office.Interop.Word.Range rng = null;

                    //MessageBox.Show("");

                    //CREATING BOOKMARK OBJECT SPECIFIC TO NOTI-UPDATE TEMPLATE
                    if (GlobalData.gNoti_Type == "N" || GlobalData.gNoti_Type == "U")
                    {

                        object obkActualImpact = "bkActualImpact";
                        rng = gendoc.Bookmarks.get_Item(ref obkActualImpact).Range;
                        rng.Text = txtBIA.Text;
                        //MessageBox.Show("Actual Impact Bk");

                        object obkPotentialImpact = "bkPotentialImpact";
                        rng = gendoc.Bookmarks.get_Item(ref obkPotentialImpact).Range;
                        rng.Text = txtBIP.Text;
                        //MessageBox.Show("Potential Impact Bk");

                        object obkEstimatedResolution = "bkEstimatedResolutionTime";
                        rng = gendoc.Bookmarks.get_Item(ref obkEstimatedResolution).Range;
                        rng.Text = txtEstimatedResolution.Text + " " + GlobalData.gTimeZone; ;

                        object obkNextUpdate = "bkNextUpdate";
                        rng = gendoc.Bookmarks.get_Item(ref obkNextUpdate).Range;
                        rng.Text = txtNextDT.Text + " " + GlobalData.gTimeZone;

                        object obkNotiUpdate = "bkNotiUpdate";
                        rng = gendoc.Bookmarks.get_Item(ref obkNotiUpdate).Range;

                        //TITLE TEXT IN THE HEADER PORTION OF THE DOCUMENT
                        int iUpNo = GlobalData.gUpNo;

                        if (iUpNo > 0)
                        {
                           rng.Text = GlobalData.gType + " " + iUpNo;//+ " - "; //Changed according to the new template +ApplicationName;
                        }

                        else
                        {
                            rng.Text = GlobalData.gType;
                        }
                    }



                    //CREATING BOOKMARKS SPECIFIC TO THE FINAL TEMPLATE
                    else
                    {
                        object obkImpact = "bkImpact";
                        rng = gendoc.Bookmarks.get_Item(ref obkImpact).Range;
                        rng.Text = txtBIA.Text;

                        object obkStrtDT = "bkStartDateTime";
                        rng = gendoc.Bookmarks.get_Item(ref obkStrtDT).Range;
                        rng.Text = txtStartDT.Text + " " + GlobalData.gTimeZone;

                        object obkResDT = "bkResolutionDateTime";
                        rng = gendoc.Bookmarks.get_Item(ref obkResDT).Range;
                        rng.Text = txtNextDT.Text + " " + GlobalData.gTimeZone;

                        //The Resolution text as per new templates.
                        object obkResolutionCommment = "bkResolutionComment";
                        rng = gendoc.Bookmarks.get_Item(ref obkResolutionCommment).Range;
                        string sLoweredSeverity = rdbSev3.Checked == true ? "3" : "4";
                        string sResolutioComment;
                        if (chkLowerSeverity.Checked == true)
                        {
                            sResolutioComment = " The Severity of this incident has been lowered from Severity '" + GlobalData.gSeverity + "' to Severity '" + sLoweredSeverity + "'.As a result work will continue to resolve the issue but here will be no further communications.";

                        }
                        else
                        {
                            sResolutioComment = "This incident has now been resolved";
                        }
                        rng.Text = sResolutioComment;

                        //Checking if its the First and Final Template
                        if (GlobalData.gNoti_Type == "FF")
                        {
                            object obkFirstAndLastNotification = "bkFirstAndLastNotification";
                            rng = gendoc.Bookmarks.get_Item(ref obkFirstAndLastNotification).Range;
                            rng.Text = "Notification/Final Communication";
                        }

                    }




                    rng = gendoc.Bookmarks.get_Item(ref obkSev).Range;
                    rng.Text = "" + GlobalData.gSeverity;

                    rng = gendoc.Bookmarks.get_Item(ref obkIncTil).Range;
                    rng.Text = "" + txtIncidentTitle.Text;

                    rng = gendoc.Bookmarks.get_Item(ref obkBrSum).Range;
                    rng.Text = "" + txtBriefSummary.Text;

                    rng = gendoc.Bookmarks.get_Item(ref obkSummP2).Range;
                    rng.Text = txtSummP1.Text;

                    rng = gendoc.Bookmarks.get_Item(ref obkSummP3).Range;
                    rng.Text = txtSummP2.Text;

                    rng = gendoc.Bookmarks.get_Item(ref obkTckRef).Range;
                    rng.Text = "" + txTicketRef.Text;

                    rng = gendoc.Bookmarks.get_Item(ref obkListApp).Range;
                    rng.Text = "" + GlobalData.gApplicatioName;

                    rng = gendoc.Bookmarks.get_Item(ref obkDetails).Range;
                    rng.Text = "" + txtDetails.Text;

                    rng = gendoc.Bookmarks.get_Item(ref obkPrevAct).Range;
                    rng.Text = "" + txtPreviousActions.Text;

                    rng = gendoc.Bookmarks.get_Item(ref obkPOC).Range;
                    rng.Text = txtPOCName.Text + "\n" + txtPOCContact.Text;

                    gendoc.SaveAs(ref filename, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                   ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                    #region :COPYING THE COMPLETE WORD DOCUMNETS CONTENTS TO BE USED AS A BODY FOR SENDING MAILS

                    gendoc.ActiveWindow.Selection.WholeStory();

                    GlobalData.gEmailBody = gendoc.ActiveWindow.Selection.Text;

                    #endregion
                    //

                    //Response.Write("<script language='javascript'>alert('Notification Created');</script>");
                    // MessageBox.Show("Word document Created Successfully .");

                }

                catch (COMException ex)
                {
                    //string msg = 1;
                    //ClientScriptManager.RegisterStartupScript(ClientScriptManager.GetType, msg, "alert ( Exception raised )", true);
                    //throw ex;
                    MessageBox.Show("There was a problem generating the notification. There seems to be a problem with the MS WORD on your system. Try re-opening the application again", "Well this is embarrassing");
                   
                }
                finally
                {
                    object saveChanges = false;
                    object originalFormat = System.Reflection.Missing.Value;
                    object routeDocument = System.Reflection.Missing.Value;

                    string msg;


                }


                /// doc.SaveAs(ref fileName, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                //             ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

                #endregion

            }

            #region: CHECKPOINT FOR CHECKING THE SANITY OF THE INFORMATION

            
            #endregion

        }

        private void lblPutTestData_Click(object sender, EventArgs e)
        {
            string sample = "This is sample incident information";
            string incident = "INC123123123";
            string sample1 = "This is sample incident information " + Environment.NewLine + "This is sample incident information." + Environment.NewLine + "This is sample incident information";

            txtIncidentTitle.Text = GlobalData.gApplicatioName + " Application is Down";
            txtBriefSummary.Text = "Users Accessing the "+GlobalData.gApplicatioName +" Application are recieving thee error Message --Error Message--  ";
            txtSummP1.Text = "Applcation team is investigating.";
            txtSummP2.Text = "This is an initial communication";
            txtDetails.Text = "Incident Started at " + txtStartDT.Text + " " + lblSTZ.Text + " " + Environment.NewLine + "" + txtBriefSummary.Text;
            txtBIA.Text = "Currently 5000 users are getting affected";
            txtBIP.Text = "10000 users use the " + GlobalData.gApplicatioName + " Application" + Environment.NewLine + "5000 users in the EMEA region are affected";
            txtPreviousActions.Text = "This is initial communication";
            
            txtPOCName.Text = "Rajesh Shetty";
            txtPOCContact.Text = "+91 9158889692";
            txtPOCEmail.Text = "rajesh.shetty@Ikran.com";

        }
               
        private void txtSend_Click(object sender, EventArgs e)
        {
            SendConfirm frm = new SendConfirm();
            frm.ShowDialog();
            frm.Dispose();
            if (GlobalData.gHasNotificationSent == 1)
            {
                this.Close();
            }
            GC.Collect();

        }

        private void lblErrors_Click(object sender, EventArgs e)
        {

        }

        private void btnTestAT_Click(object sender, EventArgs e)
        {
            //DateTime diff =  new DateTime();
            //Functions fobj = new Functions();
            //fobj.UpdateAvailabilityTracker(GlobalData.gStartDT, GlobalData.gTempNextDT, txtOutageDuration.Text, GlobalData.gPlannedActivity, GlobalData.gDueToAM, txtActionOwner.Text);


        }

        private void chkPlanned_Click(object sender, EventArgs e)
        {
            GlobalData.gPlannedActivity = chkPlanned.Checked == true ? "Y":"N";

        }

        private void chkDueToAM_Click(object sender, EventArgs e)
        {
            GlobalData.gDueToAM = chkDueToAM.Checked == true ? "Y" : "N";   
        }

        private void chkAT_Click(object sender, EventArgs e)
        {
            if (chkAT.Checked == true)
            {
                grpAvailTracker.Visible = true;
                txtOutageDuration.Text = CalculateOutage();
                GlobalData.gUpdateAT = 1;
            }
            else
            {
                grpAvailTracker.Visible = false;
                GlobalData.gUpdateAT = 0;
            }
        }

        public string CalculateOutage()
        {
            try
            {
                CultureInfo cul = new CultureInfo("en-US");
                DateTime dtStart = Convert.ToDateTime(GlobalData.gStartDT.ToString(), cul);
                DateTime dtEnd = Convert.ToDateTime(GlobalData.gTempNextDT.ToString(), cul);
                System.TimeSpan ts = dtEnd - dtStart;
                string sOutage = ts.Days.ToString() + " days " + ts.Hours.ToString() + ":" + ts.Minutes.ToString() + ":" + ts.Seconds.ToString();
                GlobalData.gOutageDuration = sOutage;
                GlobalData.gActionOwner = txtActionOwner.Text;
                return sOutage;
            }
            catch (Exception ex)
            {
                string s = "null";
                return s;
                //DO NOTHING
            }
         }

        private void txtOutageDuration_Enter(object sender, EventArgs e)
        {
            //txtOutageDuration.Text = CalculateOutage();
            //GlobalData.gOutageDuration = CalculateOutage();
            //GlobalData.gActionOwner = txtActionOwner.Text;
        }


        public void ValidateData()
        {
            //int iReturn = 0;
            GlobalData.gValidate = 0;

            if (rdbSev1.Checked == false && rdbSev2.Checked == false)
            {
                lblRDBErr.Visible = true;
                lblRDBErr.Text = "SEVERITY LEVEL PLEASE";
                lblRDBErr.ForeColor = System.Drawing.Color.Red;
                GlobalData.gValidate = 1;
            }
            else
            {
                lblRDBErr.Visible = false;
                //GlobalData.gValidate = 0;

            }
            if (txtIncidentTitle.Text == "")
            {
                txtIncidentTitle.BackColor = System.Drawing.Color.Khaki;
                txtIncidentTitle.Focus();
                txtIncidentTitle.BorderStyle = BorderStyle.Fixed3D;
                GlobalData.gValidate = 1;
            }
            else
            {
                txtIncidentTitle.BackColor = System.Drawing.SystemColors.Window;
                txtIncidentTitle.Focus();
                //GlobalData.gValidate = 0;
            }


            if (txTicketRef.Text == "" || txTicketRef.Text.Length < 10)
            {
                txTicketRef.BackColor = System.Drawing.Color.Khaki;
                txTicketRef.Focus();
                txTicketRef.BorderStyle = BorderStyle.Fixed3D;
                GlobalData.gValidate = 1;
            }
            else
            {
                txTicketRef.BackColor = System.Drawing.SystemColors.Window;
                txtBriefSummary.Focus();
                //GlobalData.gValidate = 0;
            }


            if(txtBriefSummary.Text == "")
            {
                txtBriefSummary.BackColor = System.Drawing.Color.Khaki;
                txtBriefSummary.Focus();
                txtBriefSummary.BorderStyle = BorderStyle.Fixed3D;
                GlobalData.gValidate = 1;
            }
            else
            {
                txtBriefSummary.BackColor = System.Drawing.SystemColors.Window;
                txtBriefSummary.Focus();
                //GlobalData.gValidate = 0;
            }

            if(txtSummP1.Text == "")
            {
                txtSummP1.BackColor = System.Drawing.Color.Khaki;
                txtSummP1.Focus();
                txtSummP1.BorderStyle = BorderStyle.Fixed3D;
                GlobalData.gValidate = 1;
            }
            else
            {
                txtSummP1.BackColor = System.Drawing.SystemColors.Window;
                txtSummP1.Focus();
                //GlobalData.gValidate = 0;
            }

            if(txtSummP2.Text == "")
            {
                txtSummP2.BackColor = System.Drawing.Color.Khaki;
                txtSummP2.Focus();
                txtSummP2.BorderStyle = BorderStyle.Fixed3D;
                GlobalData.gValidate = 1;
            }
            else
            {
                txtSummP2.BackColor = System.Drawing.SystemColors.Window;
                txtSummP2.Focus();
                //GlobalData.gValidate = 0;
            }

            if(txtDetails.Text == "")
            {
                txtDetails.BackColor = System.Drawing.Color.Khaki;
                txtDetails.Focus();
                txtDetails.BorderStyle = BorderStyle.Fixed3D;
                GlobalData.gValidate = 1;
            }
            else
            {
                txtDetails.BackColor = System.Drawing.SystemColors.Window;
                txtDetails.Focus();
                //GlobalData.gValidate = 0;
            }

            if (txtBIA.Text == "")
            {
                txtBIA.BackColor = System.Drawing.Color.Khaki;
                txtBIA.Focus();
                txtBIA.BorderStyle = BorderStyle.Fixed3D;
                GlobalData.gValidate = 1;
            }
            else
            {
                txtBIA.BackColor = System.Drawing.SystemColors.Window;
                txtBIA.Focus();
                //GlobalData.gValidate = 0;
            }
            if(GlobalData.gNoti_Type == "F" || GlobalData.gNoti_Type == "FF")
                 txtBIP.Text = "Sample Text To avoid a bug";
            if (txtBIP.Text == "")
            {
                txtBIP.BackColor = System.Drawing.Color.Khaki;
                txtBIP.Focus();
                txtBIP.BorderStyle = BorderStyle.Fixed3D;
                GlobalData.gValidate = 1;
            }
            else
            {
                txtBIP.BackColor = System.Drawing.SystemColors.Window;
                txtBIP.Focus();
                //GlobalData.gValidate = 0;
            }

            if (txtPreviousActions.Text == "")
            {
                txtPreviousActions.BackColor = System.Drawing.Color.Khaki;
                txtPreviousActions.Focus();
                txtPreviousActions.BorderStyle = BorderStyle.Fixed3D;
                GlobalData.gValidate = 1;
            }
            else
            {
                txtPreviousActions.BackColor = System.Drawing.SystemColors.Window;
                txtPreviousActions.Focus();
                //GlobalData.gValidate = 0;
            }

            if (txtStartDT.Text == "")
            {
                txtStartDT.BackColor = System.Drawing.Color.Khaki;
                txtStartDT.Focus();
                txtStartDT.BorderStyle = BorderStyle.Fixed3D;
                GlobalData.gValidate = 1;
            }
            else
            {
                txtStartDT.BackColor = System.Drawing.SystemColors.Window;
                txtStartDT.Focus();
                //GlobalData.gValidate = 0;
            }

            if (txtNextDT.Text == "")
            {
                txtNextDT.BackColor = System.Drawing.Color.Khaki;
                txtNextDT.Focus();
                txtNextDT.BorderStyle = BorderStyle.Fixed3D;
                GlobalData.gValidate = 1;
            }
            else
            {
                txtNextDT.BackColor = System.Drawing.SystemColors.Window;
                txtNextDT.Focus();
                //GlobalData.gValidate = 0;
            }
            if (GlobalData.gNoti_Type == "F" || GlobalData.gNoti_Type == "FF")
             txtEstimatedResolution.Text = "Sample Text To avoid a bug";
            if (txtEstimatedResolution.Text == "")
            {
                txtEstimatedResolution.BackColor = System.Drawing.Color.Aqua;
                txtEstimatedResolution.Focus();
                txtEstimatedResolution.BorderStyle = BorderStyle.Fixed3D;
                GlobalData.gValidate = 1;
            }
            else
            {
                txtEstimatedResolution.BackColor = System.Drawing.SystemColors.Window;
                txtEstimatedResolution.Focus();
                //GlobalData.gValidate = 0;
            }

            if (txtPOCName.Text == "")
            {
                txtPOCName.BackColor = System.Drawing.Color.Khaki;
                txtPOCName.Focus();
                txtPOCName.BorderStyle = BorderStyle.Fixed3D;
                GlobalData.gValidate = 1;
            }
            else
            {
                txtPOCName.BackColor = System.Drawing.SystemColors.Window;
                txtPOCName.Focus();
                //GlobalData.gValidate = 0;
            }

            if (txtPOCEmail.Text == "")
            {
                txtPOCEmail.BackColor = System.Drawing.Color.Khaki;
                txtPOCEmail.Focus();
                txtPOCEmail.BorderStyle = BorderStyle.Fixed3D;
                GlobalData.gValidate = 1;
            }
            else
            {
                txtPOCEmail.BackColor = System.Drawing.SystemColors.Window;
                txtPOCEmail.Focus();
                //GlobalData.gValidate = 0;
            }

            if (txtPOCContact.Text == "")
            {
                txtPOCContact.BackColor = System.Drawing.Color.Khaki;
                txtPOCContact.Focus();
                txtPOCContact.BorderStyle = BorderStyle.Fixed3D;
                GlobalData.gValidate = 1;
            }
            else
            {
                txtPOCContact.BackColor = System.Drawing.SystemColors.Window;
                txtPOCContact.Focus();
                //GlobalData.gValidate = 0;
            }

            if (chkAT.Checked == true)
            {

                if (txtUAStartDate.Text == "")
                {
                    txtUAStartDate.BackColor = System.Drawing.Color.Khaki;
                    txtUAStartDate.Focus();
                    txtUAStartDate.BorderStyle = BorderStyle.Fixed3D;
                    GlobalData.gValidate = 1;
                }
                else
                {
                    txtUAStartDate.BackColor = System.Drawing.SystemColors.Window;
                    txtUAStartDate.Focus();
                    //GlobalData.gValidate = 0;
                }

                if (txtUAEndDate.Text == "")
                {
                    txtUAEndDate.BackColor = System.Drawing.Color.Khaki;
                    txtUAEndDate.Focus();
                    txtUAEndDate.BorderStyle = BorderStyle.Fixed3D;
                    GlobalData.gValidate = 1;
                }
                else
                {
                    txtUAEndDate.BackColor = System.Drawing.SystemColors.Window;
                    txtUAEndDate.Focus();
                    //GlobalData.gValidate = 0;
                }

                if (txtOutageDuration.Text == "")
                {
                    txtOutageDuration.BackColor = System.Drawing.Color.Khaki;
                    txtOutageDuration.Focus();
                    txtOutageDuration.BorderStyle = BorderStyle.Fixed3D;
                    GlobalData.gValidate = 1;
                }
                else
                {
                    txtOutageDuration.BackColor = System.Drawing.SystemColors.Window;
                    txtOutageDuration.Focus();
                    //GlobalData.gValidate = 0;
                }

                

            }

            //return iReturn;
        }

        public void ValidateDate()
        {
            //strign sTempStartDT = StartTimeDT.to
            //DateTime dt;
            
            string sStartDT =  txtStartDT.Text.Trim();
            string sNextDT = txtNextDT.Text.Trim();
            string sEstDT = txtEstimatedResolution.Text.Trim();
            string sError = "";
            try 
            { DateTime dt = DateTime.Parse(sStartDT); }
            catch (Exception ex) 
            {
                //MessageBox.Show("Start Date\\Time is not entered in proper format");
                sError = " Start Date\\Time";
                GlobalData.gValidateDate = 1; 
            }

            try { DateTime dt = DateTime.Parse(sNextDT); }
            catch (Exception ex)
            {
                //MessageBox.Show("Next Date\\Time is not entered in proper format");
                sError = " Next Date\\Time";
                GlobalData.gValidateDate = 1;
            }
            //THIS IS NOT APPLICABLE WHEN IT IS A f OR AND FF.
            if(GlobalData.gNoti_Type == "F" || GlobalData.gNoti_Type == "FF")
            {
                //DO NOTHING
            }
            else
            {
            try { DateTime dt = DateTime.Parse(sEstDT); }
            catch (Exception ex)
            {
                //MessageBox.Show("Estimated Date\\Time is not entered in proper format");
                sError = " Estimated Date\\Time";
                GlobalData.gValidateDate = 1;
            }
            }
            if (GlobalData.gValidateDate == 1)
            {
                MessageBox.Show(sError + " are not in proper format ");
            }
            
            
          }

        private void lblNTZ_Click(object sender, EventArgs e)
        {

        }

        private void chkAT_CheckedChanged(object sender, EventArgs e)
        {

        }

        public void ChangeinTZ()
        {
            CultureInfo culture = new CultureInfo("en-US");
            DateTime dtTemp = Convert.ToDateTime(GlobalData.gStartDT, culture);
            txtStartDT.Text = dtTemp.ToString("dd MMMM yyyy HH:mm"); // +" " + GlobalData.gTimeZone; 


        }

        #region: FIXING THE LOCAL IMPACT WAALA HEADACHE
        private void cmbTimZone_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblNTZ.Text = cmbTimZone.Text;
            lblETZ.Text = cmbTimZone.Text;
            GlobalData.gTimeZone = cmbTimZone.Text;
        }

        
        private void chkChangeTZ_CheckedChanged(object sender, EventArgs e)
        {
             if (chkChangeTZ.Checked == true)
            {
                chkLocalTZ.Visible = true;
                cmbTimZone.Enabled = true;
            }
             else
            {
                chkLocalTZ.Visible = false;
                cmbTimZone.Enabled = false;
                //////////////////////////
                lblLocalTZ.Visible = false;
                txtLTZ.Visible = false;
                btnSetTZ.Visible = false;
                lblexTZ.Visible = false;
            }
        }

        private void btnSetTZ_Click(object sender, EventArgs e)
        {
            string TZ = txtLTZ.Text.Trim();
            cmbTimZone.Items.Add(TZ);
            GlobalData.gTimeZone = TZ;
            //cmbTimZone.SelectedText = TZ;

            chkLocalTZ.Visible = false;
            lblLocalTZ.Visible = false;
            txtLTZ.Visible = false;
            btnSetTZ.Visible = false;
            cmbTimZone.Enabled = true;
            lblexTZ.Visible = false;


            MessageBox.Show("The requested TimeZone "+ TZ +" has been added.Select the same from the dropdown and kindly modify the Start,Next and Resolutions time's accordingly.","New TimeZone Added");

        }

        private void chkLocalTZ_CheckedChanged(object sender, EventArgs e)
        {
            if (chkLocalTZ.Checked == true)
            {
                cmbTimZone.Enabled = true;
                chkLocalTZ.Visible = true;
                lblLocalTZ.Visible = true;
                txtLTZ.Visible = true;
                btnSetTZ.Visible = true;
                lblexTZ.Visible = true;
            }

            else
            {
                //chkLocalTZ.Visible = false;
                lblLocalTZ.Visible = false;
                txtLTZ.Visible = false;
                btnSetTZ.Visible = false;
                lblexTZ.Visible = false;
                //cmbTimZone.Enabled = false;
            }
        }

#endregion:  

        private void txtUAEndDate_Leave(object sender, EventArgs e)
        {
            CultureInfo cul = new CultureInfo("en-US");
            DateTime dtStart = Convert.ToDateTime(txtUAStartDate.Text, cul);
            DateTime dtEnd = Convert.ToDateTime(txtUAEndDate.Text, cul);
            System.TimeSpan ts = dtEnd - dtStart;
            string sOutage = ts.Days.ToString() + " days " + ts.Hours.ToString() + ":" + ts.Minutes.ToString() + ":" + ts.Seconds.ToString();
            txtOutageDuration.Text = sOutage;
                        
        }

        
    }
}
