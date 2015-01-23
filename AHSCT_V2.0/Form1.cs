using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Windows;
using System.Globalization;
using System.Threading;
using System.Diagnostics;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.Mail;


//using Microsoft.Office.Interop.Excel;


namespace maddytry1
{
    public partial class frmconsole : Form
    {
        public DataSet dstAplication = new DataSet();
        public DataSet dstMaster = new DataSet();
        public DataSet dstConsole = new DataSet();
        Operations obj = new Operations();
        Thread th_LoadCurrentInfo;
        //DataTable dt = new DataTable();
        System.Data.DataTable Availdt = new System.Data.DataTable();
        public int[] iDelayed = new int[50];
        public int[] iWarning = new int[50];

            

        public frmconsole()
        {
            this.Visible = false;
            //MessageBox.Show(this.Visible.ToString());
            InitializeComponent();

            #region: SETTING UP THE BACKGROUND IMAGES FOR THE APPLICATION

            //string sImageFileName = Application.StartupPath.ToString() + "\\images\\bg1.jpg";
            //pnlHeader.BackgroundImage = new Bitmap(sImageFileName);
            //pnlHeader.BackgroundImageLayout = ImageLayout.Stretch;
            //sImageFileName = Application.StartupPath.ToString() + "\\images\\bg13.jpg";
            //tbHome.BackgroundImage = new Bitmap(sImageFileName);
            //tbHome.BackgroundImageLayout = ImageLayout.Stretch;
            //tbSearch.BackgroundImage = new Bitmap(sImageFileName);
            //tbSearch.BackgroundImageLayout = ImageLayout.Stretch;


            #endregion
        }

        private void frmconsole_Load(object sender, EventArgs e)
        {
            #region: LOADING THE SPLASH SCREEN
            this.Visible = false;
            //MessageBox.Show(this.Visible.ToString());

            Splash frmSplash = new Splash();
            frmSplash.ShowDialog();


            while (GlobalData.gSplashComplete == 0)
            {
                System.Threading.Thread.Sleep(1000);
            }

                #region: EXITING THE APPLICATION IF AHSCT IS NOT ABLE TO CONNECT TO THE DATABASE

                    if (GlobalData.gUnableToConnect == 1)
                    {
                        notiAHSCT.Dispose();
                        //string source = @"D;\AHSCT_NewVersion";
                        //Directory.Exists("D:\\AHSCT_BackupData");
                        //File.Exists("D:\\AHSCT_BAckupData");
                        
                        //MessageBox.Show(Application.ProductName.ToString() + Application.ProductVersion.ToString());
                        Application.ExitThread();
                        Application.Exit();
                        
                    }
                    else
                    {
                        //do nothing
                        #region: EMAILING THE STATUS OF THE CURRENT USERS
                        try
                        {
                            string sysname = System.Environment.MachineName.ToString();
                            string uid1 = System.Environment.UserName.ToString();

                            //MessageBox.Show("System NAme : " +sysname+ " UID : " +uid + " UserName : " +uid1+ "  " + GlobalData.gCurrentUser);

                            MailMessage Send_Info = new MailMessage();
                            Send_Info.From = new MailAddress(uid1 + "@Ikran.com");
                            Send_Info.To.Add("rajesh.shetty@Ikran.com");
                            Send_Info.Subject = "AHSCT:NOTIFY I am using your tool";
                            Send_Info.Body = "System Name : " + sysname + Environment.NewLine + "User Name : " + uid1;
                            SmtpClient client = new SmtpClient("172.19.98.22", 25);
                            client.UseDefaultCredentials = true;

                            //client.Send(Send_Info);



                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error Sending Email", "Problem Sending Mail", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                        }
                        #endregion
                    }

                    this.Show();

                #endregion

             #endregion

                    lblUserDetails.Text = "WELCOME " + GlobalData.gCurrentUser;

            //MessageBox.Show(this.Height.ToString());
            //MessageBox.Show(this.Width.ToString());

            #region : ADJUSTIING THE TAB CONTROLS WIDTH AND HEIGHT AND VISBILITY OF ALL CONTROLS

            pnlHeader.Width = this.Width - 1;

            tcHome.Width = this.Width - 2;
            tcHome.Height = this.Height - 50;

            grpFavourites.Visible = false;
            //grpDropDown.Width = this.Width - 45;
            grpConsoleGrid.Width = this.Width - 45;
            grpActions.Width = this.Width - 430;
            grpConsoleGrid.Height = this.Height - 100;
            dgConsoleGrid.Width = grpConsoleGrid.Width - 20;
            //dgConsoleGrid.Height = grpConsoleGrid.Height - 5;

            //MAKING THE GENERATE BUTTONS VISIBLE 
            grpActions.Visible = false;
            //btnGenerateNew.Visible = false;
            //btnGenerateFF.Visible = false;

            btnFinal.Visible = false;
            btnNext.Visible = false;
            btnCreateSimilar.Enabled = false;


            //SETTINGS FOR THE SEARCH TAB
            dgSearchApp.Width = spltcntSearch.Panel1.Width - 20;
            dgIncidentDetails.Height = this.Height - 40;
            dgIncidentDetails.Width = this.Width - 280;
            dgIncidentDetails.Columns[1].Width = spltcntSearch.Panel2.Width - 100;
            dgSearchApp.Visible = false;
            dgIncidentDetails.Visible = false;

            lblByApp.Width = dgSearchApp.Width + 70;
            lblByIncident.Width = this.Width - 480;

            //SETTINGS FOR THE AVAILABILITY TAB
            grpATResults.Width = this.Width - 45;
            grpATResults.Height = this.Width - 10;
            grpFilter.Width = this.Width - 35;
            dgAT.Width = grpATResults.Width - 10;
            dgAT.Height = grpATResults.Height - 720;

            grpATRegions.Enabled = false;
            grpATBundles.Enabled = false;
            grpATApp.Enabled = false;
            chkByApp.Enabled = false;

            lblExportProgress.Visible = false;
            pbExportProgress.Visible = false;
            btnExportData.Visible = false;

            for (int i = 0; i < 50; i++)
            {
                iDelayed[i] = 0;                      
                iWarning[i] = 0;
            }


            #endregion  

            #region : FIXING THE TIMEZONE PROBLEM

                TimeZone local = TimeZone.CurrentTimeZone;
                ////tslblTimeZone.Text = local.StandardName + " " + local.DaylightName;

                ////LOADING GLOBAL VALUES WITH TIMEZONE VALUES
                string sTimeZone = local.StandardName;

                switch (sTimeZone)
                {
                    case "Eastern Standard Time": GlobalData.gTimeZone = "US TIME";
                        tslblTimeZone.Text = tslblTimeZone.Text + GlobalData.gTimeZone;
                        break;
                    case "GMT Standard Time": GlobalData.gTimeZone = "UK TIME";
                        tslblTimeZone.Text = tslblTimeZone.Text + GlobalData.gTimeZone;
                        break;

                    case "W. Europe Standard Time": GlobalData.gTimeZone = "SWE TIME";
                        tslblTimeZone.Text = tslblTimeZone.Text + GlobalData.gTimeZone;
                        break;

                }

                if (GlobalData.gTimeZone == "US TIME")
                {
                    stslblUSTime.Text = DateTime.Now.ToString("dd MMMM yyyy HH:mm ss") + "        ";
                    stslblUKTime.Text = DateTime.Now.AddHours(5).ToString("dd MMMM yyyy HH:mm ss") + "        ";
                    stslblSWETime.Text = DateTime.Now.AddHours(6).ToString("dd MMMM yyyy HH:mm ss") + "        ";

                }
                else if (GlobalData.gTimeZone == "UK TIME")
                {
                    stslblUKTime.Text = DateTime.Now.ToString("dd MMMM yyyy HH:mm ss") + "        ";
                    stslblUSTime.Text = DateTime.Now.AddHours(-5).ToString("dd MMMM yyyy HH:mm ss") + "        ";
                    stslblSWETime.Text = DateTime.Now.AddHours(1).ToString("dd MMMM yyyy HH:mm ss") + "        ";

                }
                else
                {
                    stslblSWETime.Text = DateTime.Now.ToString("dd MMMM yyyy HH:mm ss") + "        ";
                    stslblUSTime.Text = DateTime.Now.AddHours(-6).ToString("dd MMMM yyyy HH:mm ss") + "        ";
                    stslblUKTime.Text = DateTime.Now.AddHours(-1).ToString("dd MMMM yyyy HH:mm ss") + "        ";

                }
            #endregion

            #region : POPULATING THE CATEGORY DROPDOWN.
            //POPULATING THE CATEGORY DROPDOWN.


            //OleDbConnection conn;
            ////string connstr = "Provider=Microsoft.Jet.OleDB.4.0;Data Source =D:\\DBAS12.mdb";
            //string connstr = GlobalData.gAS12_connnectionString;
            //conn = new OleDbConnection(connstr);
            //OleDbCommand cmd = new OleDbCommand();
            ////cmd.CommandText = "Select DISTINCT AppName from Application where AppGroup = '" + cmbCategory.Text + "'";
            //cmd.CommandText = "Select DISTINCT AppGroup from Application ";
            //OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            ////OleDbDataReader dr = new OleDbDataReader(cmd);

            //da.SelectCommand.Connection = conn;
            //da.SelectCommand = cmd;
            //DataSet ds1 = new DataSet();
            //DataTable dt = new DataTable();
            //da.Fill(dt);
            //da.Fill(ds1);
            //cmbCategory.DataSource = dt;
            //cmbCategory.DisplayMember = "AppGroup";
            //cmbCategory.ValueMember = "AppGroup";

            #endregion
            
            #region : POPULATING THE CATEGORY DROPDOWN WITHOUT DATABINDING

            //cmbCategory.Items.Add("--SELECT--");
            //cmbCategory.Items.Add("Bundle-B");
            //cmbCategory.Items.Add("Bundle-D");
            //cmbCategory.Items.Add("Visage");
            //cmbCategory.Text = "--SELECT--";
            //cmbCategory.SelectedIndex = 0;

            #endregion


            #region: CHECKING IF THE USER IS USING THE OLD VERSION AND IF A NEW VERSION IS REQUIRED COPY THE SAME TO THE DESIRED ROOT FOLDER.

                if (GlobalData.gUnableToConnect == 1)
                {
                    Application.ExitThread();
                    Application.Exit();
                }
                else
                {

                    //MATCHING THE CURRENT VERSION WITH THE LATEST AVAILABLE VERSION 

                    string sNewVersionFilename = "";
                    OleDbCommand cmd = new OleDbCommand("select filename from Version", GlobalData.GlobalConnection);
                    OleDbDataReader dr = cmd.ExecuteReader();

                    while (dr.Read())
                    {
                        sNewVersionFilename = dr[0].ToString();
                        //MessageBox.Show(sNewVersionFilename);
                    }

                    if (GlobalData.gCurrentFileName == sNewVersionFilename)
                    {
                        //DO NOTHING
                    }
                    else
                    {
                        try
                        {
                            MessageBox.Show("You are not using a latest version of the application.AHSCT will upgrade itself to the latest version.Please press OK to continue.", "UPGRADE AHSCT",MessageBoxButtons.OK, MessageBoxIcon.Hand);
                            //COPYING THE DATA FROM THE <-- SOURCE --> TO <--DESTINATION-->
                            //ONE WHICH SHOWS THE PROGRESS OF THE THE COPYING PROCESS.
                            File.Copy("D:\\AHSCT_Versions\\" + sNewVersionFilename, Application.StartupPath.ToString() + "\\" + sNewVersionFilename, true);
                            MessageBox.Show("The latest version of the AHSCT viz. " + sNewVersionFilename + " is now availaible at the the location " + Application.StartupPath.ToString() + ". Please use this upgraded version for sending notifications in the future. Thanks!!!", "AHSCT UPGRADED SUCCESSFULLY",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                            Application.ExitThread();
                            Application.Exit();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("An error has occurred while upgrading the application. Please contact Rajesh.Shetty@Ikran.com for completion of the upgrade proccess","ERROR DURING UPGRADE",MessageBoxButtons.OK,MessageBoxIcon.Error);
                        }
                    }
                }
                #endregion


            #region: UPDATING DETAILS IN THE CURRENT STATUS SECTION.

            //int rowcount = 0;
            //string typ = "Final";
            //OleDbCommand curr_cmd = new OleDbCommand("select TicketReference,ApplicationName,Type,UpdateNo,Severity,IncidentTitle,DateTime1,DateTime2 FROM Console where Type <> '" + typ + "'", conn);
            //conn.Open();
            //OleDbDataReader drr = curr_cmd.ExecuteReader();

            ////COUNTING THE NO. OF ROWS TURNED.
            //while (drr.Read())
            //{rowcount++;}

            //GlobalData.gTempRowCount = rowcount;

            //drr.Close();

            //OleDbDataReader dr1 = curr_cmd.ExecuteReader();
            //dg_CurrentInfo.Width = this.Width - 10;
            //dg_CurrentInfo.Rows.Add(50);
            //int rowc = 0;
            //while (dr1.Read())
            //{


            //    //MessageBox.Show(dr1[0].ToString());
            //    for (int i = 0; i < 8; i++)
            //    {
            //        dg_CurrentInfo.Rows[rowc].Cells[i].Value = dr1[i].ToString();

            //    }
            //    //    dg_CurrentInfo.Rows[rowc].Cells[8].Value = GetRemainingTime(dr1[7].ToString());

            //    rowc++;

            //}

            dg_CurrentInfo.Width = this.Width - 10;
            dg_CurrentInfo.Height = this.Height - 10;
            dg_CurrentInfo.Columns[8].Width = dg_CurrentInfo.Width;
            dg_CurrentInfo.BorderStyle = BorderStyle.None;

            #endregion


            #region: USING THREADING TO UPDATE THE DETAILS IN THE CURRENT STATUS SECTION


            Operations obj_Ops = new Operations(dg_CurrentInfo);
            th_LoadCurrentInfo = new Thread(new ThreadStart(obj_Ops.LoadCurrentStatus));
            th_LoadCurrentInfo.Name = "LoadCurrentInfo";
            th_LoadCurrentInfo.Start();


            #endregion


            
           
        }

        #region:  NO NEED FOR THE DROP DOWN (AS IT ACCOUNTED FOR A BAD DESIGN. LEAAARRRRNNNN )

        //private void cmbCategory_SelectedIndexChanged(object sender, EventArgs e)
        //{
        ////string test = "Bundle-B";
        //OleDbConnection conn;
        ////string connstr = "Provider=Microsoft.Jet.OleDB.4.0;Data Source =D:\\DBAS12.mdb";
        //string connstr = GlobalData.gAS12_connnectionString;
        //conn = new OleDbConnection(connstr);
        //OleDbCommand cmd = new OleDbCommand();
        //cmd.CommandText = "Select DISTINCT AppName from Application where AppGroup = '" + cmbCategory.Text + "'";
        //cmd.CommandText = "Select * from Application ";
        //OleDbDataAdapter da = new OleDbDataAdapter(cmd);
        //OleDbDataReader dr = new OleDbDataReader(cmd);

        //da.SelectCommand.Connection = GlobalData.GlobalConnection;
        //da.SelectCommand = cmd;
        //DataSet ds1 = new DataSet();
        //DataTable dt = new DataTable();
        //da.Fill(dt);
        //da.Fill(ds1);

        //cmbApplication.DataSource = dt;



        //OleDbConnection conn = new OleDbConnection(GlobalData.gAS12_connnectionString);
        //OleDbCommand cmd = new OleDbCommand("Select * from Application");
        //OleDbDataAdapter dadp = new OleDbDataAdapter();
        //dadp.SelectCommand = cmd;
        //dadp.SelectCommand.Connection = conn;
        //dadp.Fill(GlobalData.dstAplication);


        //    string sCategory;

        //    sCategory = cmbCategory.Text == "BUNDLE-B" ? "B" : "D";


        //    DataView dv = GlobalData.dstAplication.Tables[0].DefaultView;
        //    dv.RowFilter = "AppGroup='"+sCategory+"'";

        //    //cmbApplication.Items.Clear();
        //    cmbApplication.DataSource = dv;
        //    cmbApplication.DisplayMember = "AppName";
        //    cmbApplication.ValueMember = "AppName";


        //    GlobalData.gCategory = cmbCategory.Text;

        //    btnFinal.Visible = false;
        //    btnNext.Visible = false;

        //}


        #endregion

        private void cmbApplication_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                //MAKING THE GENERATE BUTTONS VISIBLE 
                grpActions.Visible = true;
                //btnGenerateNew.Visible = true;
                //btnGenerateFF.Visible = true;

                string type = "Final";
                string appname = cmbApplication.Text;
                GlobalData.gApplicatioName = cmbApplication.Text;
                //OleDbConnection conn;
                //string connstr = GlobalData.gAS12_connnectionString;



                //conn = new OleDbConnection(connstr);
                OleDbCommand cmd = new OleDbCommand();
                cmd.CommandText = "SELECT [TicketReference] AS INCIDENT_NO,[ApplicationName] AS APPLICATION_NAME,[Type] AS TYPE,[UpdateNo] AS UPDATE_NO,[Severity] AS SEVERITY,[IncidentTitle] AS TITLE,[DateTime1] AS START_DATE_TIME,[DateTime2] AS NEXT_UPDATE_DATE_TIME,[POC] AS POC FROM Console WHERE ((ApplicationName = '" + appname + "') AND (Type <> '" + type + "'  ))";
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.SelectCommand.Connection = GlobalData.GlobalConnection;//conn;
                da.SelectCommand = cmd;
                //DataSet ds1 = new DataSet();
                DataTable dt = new DataTable();
                da.Fill(dt);

                dgConsoleGrid.DataSource = dt;

                //GETTING THE NO. OF ROWS RETURNED BY THE QUERY
                int iRowsReturned = dt.Rows.Count;

                lblRowsReturned.Text = iRowsReturned == 0 ? "NO HIGH SEVERITY NOTIFICATIONS ARE BEING SENT FOR '" + GlobalData.gApplicatioName + "' APPLICATION" : "CURRENTLY " + iRowsReturned.ToString() + " NOTIFICATION(S) ARE BEING SENT OUT FOR '" + GlobalData.gApplicatioName + "' APPLICATION";
                lblRowsReturned.Width = this.Width - 30;
                dgConsoleGrid.Visible = iRowsReturned == 0 ? false : true;
                grpConsoleGrid.Visible = iRowsReturned == 0 ? false : true;
                grpConsoleGrid.Text = lblRowsReturned.Text;
                lblRowsReturned.Visible = iRowsReturned == 0 ? true : false;

                //conn.Close();

                btnFinal.Visible = false;
                btnNext.Visible = false;
            }
            catch (Exception ex)
            {
                Functions fobj = new Functions();
                fobj.ErrorReporting(ex.Message.ToString(), "cmbApplication_SelectedIndez", "413");
                
            }

        }

        private void btnGenerateNew_Click(object sender, EventArgs e)
        {
            //Generating a new notifcation

            btnFinal.Enabled = false;
            btnNext.Enabled = false;

            GlobalData.gNoti_Type = "N";
            GlobalData.gStatus = "Notification";
            GlobalData.gType = "Notification";
            GlobalData.gCategory = cmbCategory.Text;
            GlobalData.gApplicatioName = cmbApplication.Text;
            GlobalData.gUpNo = 0;

            //REMOVING THE BELOW SNIPPET AS IT grpBox WAS COUSING SOME VISIBILITY ISSUES.
            GlobalData.gApplicatioName = cmbApplication.Text; // pre-requisite info
            GlobalData.gCategory = cmbCategory.Text; // pre-requisite info
            Favourites oFav = new Favourites();
            oFav.ShowDialog();

            //grpFavourites.Visible = true;

            //PopulateFavourites();

            //Generate gObj = new Generate();
            //gObj.ShowDialog();
            //gObj.Dispose();

        }

        private void btnGenerateFF_Click(object sender, EventArgs e)
        {
            btnFinal.Enabled = false;
            btnNext.Enabled = false;

            GlobalData.gNoti_Type = "FF";
            GlobalData.gStatus = "Notification/Final";
            GlobalData.gType = "Final";
            GlobalData.gCategory = cmbCategory.Text;
            GlobalData.gApplicatioName = cmbApplication.Text;
            GlobalData.gUpNo = 0;

            //grpFavourites.Visible = true;

            //PopulateFavourites();

            Generate gObj = new Generate();
            gObj.Show();
            //gObj.Dispose();

        }

        private void dgConsoleGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            btnFinal.Enabled = true;
            btnFinal.Visible = true;
            btnNext.Visible = true;
            btnNext.Enabled = true;
            GlobalData.gCurrTicketRef = dgConsoleGrid.CurrentRow.Cells[0].Value.ToString();
            //MessageBox.Show(GlobalData.gCurrTicketRef);

        }

        private void btnNext_Click(object sender, EventArgs e)
        {

            GlobalData.gNoti_Type = "U";
            GlobalData.gStatus = "Update";
            GlobalData.gType = "Update";

            string CurrentTicket = GlobalData.gCurrTicketRef;
            Functions obj = new Functions();
            obj.GetUpdateDetails(CurrentTicket);
            GlobalData.gUpNo++;
            Generate frm = new Generate();
            frm.Text = GlobalData.gStatus + " " + GlobalData.gUpNo + " - " + GlobalData.gApplicatioName;
            frm.ShowDialog();
            frm.Dispose();

        }

        private void btnFinal_Click(object sender, EventArgs e)
        {
            GlobalData.gNoti_Type = "F";
            GlobalData.gStatus = "Final Communication";
            GlobalData.gType = "Final";

            string CurrentTicket = GlobalData.gCurrTicketRef;
            Functions obj = new Functions();
            obj.GetUpdateDetails(CurrentTicket);
            obj = null;
            GlobalData.gUpNo++;
            Generate frm = new Generate();
            frm.Text = GlobalData.gStatus + " " + GlobalData.gUpNo + " - " + GlobalData.gApplicatioName;
            frm.ShowDialog();
            frm.Dispose();
        }

        private void dgConsoleGrid_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            //MessageBox.Show("Complete");
            if (dgConsoleGrid.Rows.Count == 0) ;
            {
                //MessageBox.Show("Currently no Hight Severity Notifications are being Sent Out for the " + GlobalData.gApplicatioName + "Aplication");
            }

        }

        private void lblRowsReturned_Click(object sender, EventArgs e)
        {

        }

        private void btnCreateSimilar_Click(object sender, EventArgs e)
        {

            string CurrentTicket = GlobalData.gCurrTicketRef;
            Functions obj = new Functions();
            obj.GetUpdateDetails(CurrentTicket);
            GlobalData.gNoti_Type = "FN";
            GlobalData.gUpNo = 0;
            GlobalData.gCurrTicketRef = "";
            Generate frm = new Generate();
            frm.Text = GlobalData.gStatus + " " + GlobalData.gUpNo + " - " + GlobalData.gApplicatioName;
            frm.ShowDialog();
            frm.Dispose();
        }


        public void PopulateFavourites()
        {
            try
            {
                //THE FUNCTION WILL POPULATE THE FAVOURTES GRID 
                string sType = "Notification";
                //OleDbConnection conn;
                //string connstr = "Provider=Microsoft.Jet.OleDB.4.0;Data Source =D:\\DBAS12.mdb";
                //string connstr = GlobalData.gAS12_connnectionString;
                //conn = new OleDbConnection(connstr);
                OleDbCommand cmd = new OleDbCommand();
                GlobalData.gApplicatioName = cmbApplication.Text;
                string a = "Notification";
                cmd.CommandText = "Select DISTINCT [TicketReference] AS INCIDENT_NO,[IncidentTitle] AS INCIDENT_TITLE,[ApplicationName] AS APPLICATION_NAME,[Severity] AS SEVERITY,[DateTime1] AS START_DATE  FROM Master WHERE ApplicationName = '" + GlobalData.gApplicatioName + "' AND Type = '" + a + "'";//where AppGroup = '" + cmbCategory.Text + "'";            //cmd.CommandText = "Select TicketReference,IncidentTitle From Master Where ApplicationName = "+GlobalData.gApplicatioName +"";//AS INCIDENT_TITLE,[ApplicationName] AS APPLICATION_NAME,[Severity] AS SEVERITY,[DateTime1] AS START_DATE  FROM Console WHERE TicketReference IN (SELECT TicketReference from console where ApplicationName = "+GlobalData.gApplicatioName+")"; 
                //cmd.CommandText = "Select DISTINCT(AppName,AppGroup) from Application ";

                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.SelectCommand.Connection = GlobalData.GlobalConnection; //conn;
                da.SelectCommand = cmd;
                DataTable dt = new DataTable();
                da.Fill(dt);

                //CHECKING IF THERE ARE ANY FAVOURITES THAT EXIST FOR THIS PARTICULAR APPLICATION.
                int iCheckRowsReturned = dt.Rows.Count;

                if (iCheckRowsReturned == 0)
                {

                    grpFavourites.Visible = false;

                    //MessageBox.Show(" NO FAVOURITES AVAILAIBLE FOR THIS APPLICATION");

                    //Generating a new notifcation

                    btnFinal.Enabled = false;
                    btnNext.Enabled = false;

                    GlobalData.gNoti_Type = "N";
                    GlobalData.gStatus = "Notification";
                    GlobalData.gType = "Notification";
                    GlobalData.gCategory = cmbCategory.Text;
                    GlobalData.gApplicatioName = cmbApplication.Text;
                    GlobalData.gUpNo = 0;

                    Generate gObj = new Generate();
                    gObj.ShowDialog();
                    gObj.Dispose();


                }
                else
                {
                    dgFavourites.DataSource = dt;
                }
            }
            catch (Exception ex)
            {
                Functions fobj = new Functions();
                fobj.ErrorReporting(ex.Message.ToString(),"PopulateFavourites", "596");
            }

        }

        private void dgFavourites_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            btnCreateSimilar.Enabled = true;
            GlobalData.gCurrTicketRef = dgFavourites.CurrentRow.Cells[0].Value.ToString();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            grpFavourites.Visible = false;
        }

        private void btnCreateNew_Click(object sender, EventArgs e)
        {
            //Generating a new notifcation

            btnFinal.Enabled = false;
            btnNext.Enabled = false;

            GlobalData.gNoti_Type = "N";
            GlobalData.gStatus = "Notification";
            GlobalData.gType = "Notification";
            GlobalData.gCategory = cmbCategory.Text;
            GlobalData.gApplicatioName = cmbApplication.Text;
            GlobalData.gUpNo = 0;

            Generate gObj = new Generate();
            gObj.ShowDialog();
            gObj.Dispose();

        }

        private void btnGenerateNew_MouseHover(object sender, EventArgs e)
        {
            //btnGenerateNew.ForeColor = System.Drawing.Color.White;
            // btnGenerateNew.BackColor = System.Drawing.Color.LightGray;

        }

        private void btnGenerateNew_MouseLeave(object sender, EventArgs e)
        {
            //btnGenerateNew.BackColor = System.Drawing.Color.White;
        }

        private void btnGenerateFF_MouseHover(object sender, EventArgs e)
        {
            //btnGenerateFF.BackColor = System.Drawing.Color.LightGray;
        }

        private void btnGenerateFF_MouseLeave(object sender, EventArgs e)
        {
            //btnGenerateFF.BackColor = System.Drawing.Color.White;
        }

        private void tmr_TIME_Tick(object sender, EventArgs e)
        {
            #region : DISPLAYING THE DATE TIME VALUES FOR THE USER FROM DIFFERENT TIMEZONE'S 
                //MessageBox.Show(GlobalData.gTimeZone.ToString());
                //tss_UKTIME.Text = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();
                //tss_US_TIME.Text = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString();

                if (GlobalData.gTimeZone == "US TIME")
                {
                    stslblUSTime.Text = DateTime.Now.ToString("dd MMMM yyyy HH:mm ss") + "        ";

                    stslblUKTime.Text = DateTime.Now.AddHours(5).ToString("dd MMMM yyyy HH:mm ss") + "        ";
                    stslblSWETime.Text = DateTime.Now.AddHours(6).ToString("dd MMMM yyyy HH:mm ss") + "        ";

                }
                else if (GlobalData.gTimeZone == "UK TIME")
                {
                    stslblUKTime.Text = DateTime.Now.ToString("dd MMMM yyyy HH:mm ss") + "        ";

                    stslblUSTime.Text = DateTime.Now.AddHours(-5).ToString("dd MMMM yyyy HH:mm ss") + "        ";
                    stslblSWETime.Text = DateTime.Now.AddHours(1).ToString("dd MMMM yyyy HH:mm ss") + "        ";

                }
                else
                {
                    stslblSWETime.Text = DateTime.Now.ToString("dd MMMM yyyy HH:mm ss") + "        ";

                    stslblUSTime.Text = DateTime.Now.AddHours(-6).ToString("dd MMMM yyyy HH:mm ss") + "        ";
                    stslblUKTime.Text = DateTime.Now.AddHours(-1).ToString("dd MMMM yyyy HH:mm ss") + "        ";

                }

            #endregion

            #region : SHOWING AVAILAVILITY EXPORT PROGRESS 
               
                if (GlobalData.gIfExportWorking == 1)
                {

                    lblExportProgress.Text = GlobalData.gExportProgressStatus + "% COMPLETED";
                }

                else
                {
                    // NO NOTHING AS NO DAT IS BEING EXPORTED
                }
                #endregion

            #region: GET THE REMAINING TIME FOR EACH NOTIFICATION SENT
             
                  GetRemainingTime();

                #endregion

            //CLOSING THE DATABASE CONNECTION 

            //conn.Close();

            ////////////////////////////////


        }

        private void button1_Click(object sender, EventArgs e)
        {
            DateTime t = DateTime.Now;
            //string s = DateTime.Now.Day.ToString() + " " + (t.ToString("MMMM")).ToUpper() +" " + DateTime.Now.Year.ToString() + "" +t.ToString("HH") + ":" + t.ToString("mm") + " UK TIME" ;
            string s = t.ToString("dd MMMM yyyy HH:MM ");// +"  UK TIME ";
            MessageBox.Show(s);


            CultureInfo culture = new CultureInfo("en-US");
            DateTime dt1 = Convert.ToDateTime(s, culture);

            MessageBox.Show(dt1.ToString());



        }


        public void GetRemainingTime()
        {

            //CALCULATING THE DATA ON RUNTIME

            int count = GlobalData.gTempRowCount;
            
            for (int i = 0; i < count; i++)
            {
                string nxt = dg_CurrentInfo.Rows[i].Cells[7].Value.ToString();
                CultureInfo culture = new CultureInfo("en-US");
                DateTime NxtUp = Convert.ToDateTime(nxt, culture);
                DateTime Curr = DateTime.Now;
                System.TimeSpan ts = NxtUp - Curr;
                int timMin = ts.Minutes;
                int timHrs = ts.Hours;
                if (timHrs > -2)
                {
                    if (timMin < 0 || ts.Hours < 0)
                    {

                        if (timMin > -12)
                        {
                            if (iDelayed[i] == 0) // NO NEED TO SEND THE MAIL IF IT HAS BEEN ALREADY SENT.
                            {
                                iDelayed[i] = 1;
                                dg_CurrentInfo.Rows[i].Cells[9].Style.ForeColor = System.Drawing.Color.Red;

                                #region: AN PRIORITY MAIL WILL BE SENT TO THE  SUPPORT TEAM STATING THEY ARE ALREADY LATE

                                try
                                {
                                    string message = "Hi," + Environment.NewLine + "You have crossed the expected time for sending the Notification/Update/Final Communication for " + dg_CurrentInfo.Rows[i].Cells[1].ToString() + " for the incident no " + dg_CurrentInfo.Rows[i].Cells[0].ToString() + ". As this has already been delayed please consider this on priority." + Environment.NewLine + Environment.NewLine + " Regards," + Environment.NewLine + "AHSCT Admin";
                                    MailMessage Send_Info = new MailMessage();
                                    Send_Info.From = new MailAddress("AHSCT_Admin@Ikran.com");
                                    Send_Info.To.Add(dg_CurrentInfo.Rows[i].Cells[8].ToString());
                                    Send_Info.CC.Add("rajesh.shetty@Ikran.com");
                                    Send_Info.Subject = "AHSCT:NOTIFY COMMUNICATION DELAYED";
                                    Send_Info.Body = message;
                                    SmtpClient client = new SmtpClient("172.19.98.22", 25);
                                    client.UseDefaultCredentials = true;
                                    //client.Send(Send_Info);

                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("Error Sending Mail", "Problem Sending Mail", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            else
                            {//DO NOTHING
                            }
                        }

                                #endregion

                    }
                    else if (timMin < 10 && timMin > 5)
                    {

                        if (iWarning[i] == 0)
                        {
                            iWarning[i] = 1;
                            #region: AN WARNING MAIL HAS TO BE SENT TO SUPPORT TEAM STATING THAT 10 MINS ARE LEFT FOR THE NEXT UPDATE

                            try
                            {
                                string message = "Hi," + Environment.NewLine + "The Notification/Update/Final Communication for " + dg_CurrentInfo.Rows[i].Cells[1].ToString() + " with the incident no " + dg_CurrentInfo.Rows[i].Cells[0].ToString() + " is due in " + dg_CurrentInfo.Rows[i].Cells[9].ToString() + Environment.NewLine + Environment.NewLine + " Regards," + Environment.NewLine + "AHSCT Admin";
                                MailMessage Send_Info = new MailMessage();
                                Send_Info.From = new MailAddress("AHSCT_Admin@Ikran.com");
                                Send_Info.To.Add(dg_CurrentInfo.Rows[i].Cells[8].ToString());
                                Send_Info.CC.Add("rajesh.shetty@Ikran.com");
                                Send_Info.Subject = "AHSCT:NOTIFY COMMUNICATION DUE IN " + dg_CurrentInfo.Rows[i].Cells[9].ToString();
                                Send_Info.Body = message;
                                SmtpClient client = new SmtpClient("172.19.98.22", 25);
                                client.UseDefaultCredentials = true;
                                //client.Send(Send_Info);

                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Error Sending Mail", "Problem Sending Mail", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            #endregion
                        }

                    }

                    else if (timMin < 5)
                    {
                        if (timMin > 3)
                        {
                            if (iWarning[i] == 0)
                            {
                                iWarning[i] = 1;
                                #region: AN WARNING MAIL HAS TO BE SENT TO SUPPORT TEAM STATING THAT 10 MINS ARE LEFT FOR THE NEXT UPDATE

                                try
                                {
                                    string message = "Hi," + Environment.NewLine + "The Notification/Update/Final Communication for " + dg_CurrentInfo.Rows[i].Cells[1].ToString() + " with the incident no " + dg_CurrentInfo.Rows[i].Cells[0].ToString() + " is due in " + dg_CurrentInfo.Rows[i].Cells[9].ToString() + Environment.NewLine + Environment.NewLine + " Regards," + Environment.NewLine + "AHSCT Admin";
                                    MailMessage Send_Info = new MailMessage();
                                    Send_Info.From = new MailAddress("AHSCT_Admin@Ikran.com");
                                    Send_Info.To.Add(dg_CurrentInfo.Rows[i].Cells[8].ToString());
                                    Send_Info.CC.Add("rajesh.shetty@Ikran.com");
                                    Send_Info.Subject = "AHSCT:NOTIFY COMMUNICATION DUE IN " + dg_CurrentInfo.Rows[i].Cells[9].ToString();
                                    Send_Info.Body = message;
                                    SmtpClient client = new SmtpClient("172.19.98.22", 25);
                                    client.UseDefaultCredentials = true;
                                    //client.Send(Send_Info);

                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show("Error Sending Mail", "Problem Sending Mail", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                                #endregion
                            }
                        }
                    }
                }
                string timerem = ts.Hours.ToString() + ":" + ts.Minutes.ToString() + ":" + ts.Seconds.ToString();
                dg_CurrentInfo.Rows[i].Cells[9].Value = timerem;
            }
        }

        private void btnSearchApp_Click(object sender, EventArgs e)
        {

            try
            {
                dgSearchApp.Visible = true;
                //POPULATING THE DATAGRID FOR APPLICATION SEARCH

                //string connstr = GlobalData.gAS12_connnectionString;
                //OleDbConnection conn = new OleDbConnection();
                //conn = new OleDbConnection(connstr);
                //OleDbCommand cmd = new OleDbCommand("select AppGroup,AppName,AppRemedyName,AppDescription,ServicelLevel,CountryOfSupport,ClientArea,GXP,SOX,UserBase,L2SDCo,L3CHCo,ChargingCategory,ASM,ONL2Manager,ONL3Manager,OFL2Manager,OFL3Manager from Application where AppName = '" + cmb_S_AppName.Text + "'", GlobalData.GlobalConnection);
                //conn.Open()
                OleDbCommand cmd = new OleDbCommand("select AppGroup,AppName,AppRemedyName,ServiceLevel,CountryOfSupport,ClientArea,GXP,SOX,UserBase,L2SDCo,L3CHCo,ChargingCategory,ASM,ONL2Manager,ONL3Manager,OFL2Manager,OFL3Manager from Application where AppName = '" + cmb_S_AppName.Text + "'", GlobalData.GlobalConnection);
                OleDbDataReader dr = cmd.ExecuteReader();

                dr.Read();
                dgSearchApp.Rows.Clear();
                dgSearchApp.Rows.Add(19);
                dgSearchApp.Rows[0].Cells[0].Value = "APPLCIATION GROUP";
                dgSearchApp.Rows[0].Cells[1].Value = dr[0].ToString();

                dgSearchApp.Rows[1].Cells[0].Value = "APPLCIATION NAME ";
                dgSearchApp.Rows[1].Cells[1].Value = dr[1].ToString();

                dgSearchApp.Rows[2].Cells[0].Value = "APPLCIATION REMEDY NAME ";
                dgSearchApp.Rows[2].Cells[1].Value = dr[2].ToString();

                dgSearchApp.Rows[3].Cells[0].Value = "SERVICE LEVEL";
                dgSearchApp.Rows[3].Cells[1].Value = dr[3].ToString();

                dgSearchApp.Rows[4].Cells[0].Value = "COUNTRY OF SUPPORT";
                dgSearchApp.Rows[4].Cells[1].Value = dr[4].ToString();

                dgSearchApp.Rows[5].Cells[0].Value = "CLIENT AREA";
                dgSearchApp.Rows[5].Cells[1].Value = dr[5].ToString();

                dgSearchApp.Rows[6].Cells[0].Value = "GXP";
                dgSearchApp.Rows[6].Cells[1].Value = dr[6].ToString();


                dgSearchApp.Rows[7].Cells[0].Value = "SOX";
                dgSearchApp.Rows[7].Cells[1].Value = dr[7].ToString();

                dgSearchApp.Rows[8].Cells[0].Value = "USER BASE";
                dgSearchApp.Rows[8].Cells[1].Value = dr[8].ToString();

                dgSearchApp.Rows[9].Cells[0].Value = "L2 ServiceDesk Co";
                dgSearchApp.Rows[9].Cells[1].Value = dr[9].ToString();

                dgSearchApp.Rows[10].Cells[0].Value = "L3 Change Co";
                dgSearchApp.Rows[10].Cells[1].Value = dr[10].ToString();

                dgSearchApp.Rows[11].Cells[0].Value = "CHARGING CATEGORY";
                dgSearchApp.Rows[11].Cells[1].Value = dr[11].ToString();

                dgSearchApp.Rows[12].Cells[0].Value = "ASM";
                dgSearchApp.Rows[12].Cells[1].Value = dr[12].ToString();

                dgSearchApp.Rows[13].Cells[0].Value = "ON L2 MANAGER";
                dgSearchApp.Rows[13].Cells[1].Value = dr[13].ToString();

                dgSearchApp.Rows[14].Cells[0].Value = "ON L3 MANAGER";
                dgSearchApp.Rows[14].Cells[1].Value = dr[14].ToString();

                dgSearchApp.Rows[15].Cells[0].Value = "OFF L2 MANAGER";
                dgSearchApp.Rows[15].Cells[1].Value = dr[15].ToString();

                dgSearchApp.Rows[16].Cells[0].Value = "OFF L3 MANAGER";
                dgSearchApp.Rows[16].Cells[1].Value = dr[16].ToString();




            }
            catch (Exception ex)
            {
                throw ex;
                lblErr.Visible = true;
                //lblErr.Text = ex.Message.ToString();
                lblErr.Text = "No Information exists for " + cmb_S_AppName.Text.ToUpper() + " application";
                dgSearchApp.Visible = false;

            }
            //}

            //else
            //{
            //MessageBox.Show("APPLICATION NOT FOUND");
            //}
            // }
        }

        private void tsCurrent_Enter(object sender, EventArgs e)
        {


            //th_LoadCurrentInfo.Abort();
            //th_LoadCurrentInfo.Start();
            //MessageBox.Show("DONE");

            if (GlobalData.gUpdateCurrentInfo == 1)
            {
                GlobalData.gUpdateCurrentInfo = 0;
                Operations obj_Ops = new Operations(dg_CurrentInfo);
                th_LoadCurrentInfo = new Thread(new ThreadStart(obj_Ops.LoadCurrentStatus));
                //th_LoadCurrentInfo.Name = "LoadCurrentInfo";
                th_LoadCurrentInfo.Start();
            }
            else
            {
                //MessageBox.Show(th_LoadCurrentInfo.ThreadState.ToString());
            }
        }

        private void tbSearch_Click(object sender, EventArgs e)
        {

        }

        private void btnSearchInc_Click(object sender, EventArgs e)
        {
            

            if (txtIncidentId.Text == "" || txtIncidentId.Text == "-- INCIDENT NO --")
            {
                MessageBox.Show("PLEASE ENTER AN INCIDENT ID","You are missing some Information");
            }
            else
            {
                try
                {
                    dgIncidentDetails.Visible = true;
                    //POPULATING THE DATAGRID FOR APPLICATION SEARCH

                    //string connstr = GlobalData.gAS12_connnectionString;
                    //OleDbConnection conn = new OleDbConnection();
                    //conn = new OleDbConnection(connstr);
                    OleDbCommand cmd = new OleDbCommand("select TicketReference,ApplicationName,Type,UpdateNo,Severity,IncidentTitle,BriefSummary,SummaryPart1,SummaryPart2,Details,BIActual,BIPotential,PreviousActions,DateTime1,DateTime2,POC from Console where TicketReference = '" + txtIncidentId.Text + "'", GlobalData.GlobalConnection);
                    //OleDbCommand cmd = new OleDbCommand("select TicketReference,ApplicationName,Type,UpdateNo,Severity from Console where TicketReference = '" + txtIncidentId.Text + "' ORDER BY UpdateNo", GlobalData.GlobalConnection);
                    //conn.Open();

                    OleDbDataReader dr = cmd.ExecuteReader();

                    dr.Read();

                    dgIncidentDetails.Rows.Add(16);
                    dgIncidentDetails.Rows[0].Cells[0].Value = "INCIDENT_ID";
                    dgIncidentDetails.Rows[0].Cells[1].Value = dr[0].ToString();

                    dgIncidentDetails.Rows[1].Cells[0].Value = "APPLICATION_NAME";
                    dgIncidentDetails.Rows[1].Cells[1].Value = dr[1].ToString();

                    dgIncidentDetails.Rows[2].Cells[0].Value = "TYPE";
                    dgIncidentDetails.Rows[2].Cells[1].Value = dr[2].ToString();

                    dgIncidentDetails.Rows[3].Cells[0].Value = "UPDATE_NO";
                    dgIncidentDetails.Rows[3].Cells[1].Value = dr[3].ToString();

                    dgIncidentDetails.Rows[4].Cells[0].Value = "SEVERITY";
                    dgIncidentDetails.Rows[4].Cells[1].Value = dr[4].ToString();

                    dgIncidentDetails.Rows[5].Cells[0].Value = "INCIDENT TITLE";
                    dgIncidentDetails.Rows[5].Cells[1].Value = dr[5].ToString();
                    //dgIncidentDetails.Rows[5].Height = 50;

                    dgIncidentDetails.Rows[6].Cells[0].Value = "BRIEF SUMMARY";
                    dgIncidentDetails.Rows[6].Cells[1].Value = dr[6].ToString();
                    dgIncidentDetails.Rows[6].Height = 50;

                    dgIncidentDetails.Rows[7].Cells[0].Value = "LAST UPDATE";
                    dgIncidentDetails.Rows[7].Cells[1].Value = dr[7].ToString();
                    dgIncidentDetails.Rows[7].Height = 50;

                    dgIncidentDetails.Rows[8].Cells[0].Value = "ACTIONS SINCE LAST UPDATE";
                    dgIncidentDetails.Rows[8].Cells[1].Value = dr[8].ToString();
                    dgIncidentDetails.Rows[8].Height = 50;

                    dgIncidentDetails.Rows[9].Cells[0].Value = "DETAILS";
                    dgIncidentDetails.Rows[9].Cells[1].Value = dr[9].ToString();
                    dgIncidentDetails.Rows[9].Height = 50;

                    dgIncidentDetails.Rows[10].Cells[0].Value = "ACTUAL IMPACT";
                    dgIncidentDetails.Rows[10].Cells[1].Value = dr[10].ToString();
                    dgIncidentDetails.Rows[10].Height = 50;

                    dgIncidentDetails.Rows[11].Cells[0].Value = "POTENTIAL IMAPACT";
                    dgIncidentDetails.Rows[11].Cells[1].Value = dr[11].ToString();
                    dgIncidentDetails.Rows[11].Height = 50;

                    dgIncidentDetails.Rows[12].Cells[0].Value = "PREVIOUS ACTIONS";
                    dgIncidentDetails.Rows[12].Cells[1].Value = dr[12].ToString();
                    dgIncidentDetails.Rows[12].Height = 50;

                    dgIncidentDetails.Rows[13].Cells[0].Value = "START DATE";
                    dgIncidentDetails.Rows[13].Cells[1].Value = dr[13].ToString();

                    dgIncidentDetails.Rows[14].Cells[0].Value = "LAST UPDATE DATE-TIME";
                    dgIncidentDetails.Rows[14].Cells[1].Value = dr[14].ToString();

                    dgIncidentDetails.Rows[15].Cells[0].Value = "POC";
                    dgIncidentDetails.Rows[15].Cells[1].Value = dr[15].ToString();
                    //dgIncidentDetails.Rows[15].Height = 50;
                    //}
                    //else
                    //{
                    //MessageBox.Show("APPLICATION NOT FOUND");
                    //}
                }

                catch (Exception ex)
                {
                    //throw ex;
                    lblIncErr.Text = "No Information is available for the Incident " + txtIncidentId.Text + ".";
                    dgIncidentDetails.Visible = false;
                }

            }

        }

        private void txtIncidentId_Enter(object sender, EventArgs e)
        {
            if (txtIncidentId.Text == "-- INCIDENT NO --")
            {
                txtIncidentId.Text = "";
            }
            txtIncidentId.ForeColor = System.Drawing.Color.Black;

        }

        private void txtIncidentId_Leave(object sender, EventArgs e)
        {
            if (txtIncidentId.Text == "")
            {
                txtIncidentId.Text = "-- INCIDENT NO --";
                txtIncidentId.ForeColor = System.Drawing.Color.Gray;
            }
        }

        private void txtAppName_Leave(object sender, EventArgs e)
        {
            if (txtAppName.Text == "")
            {
                txtAppName.ForeColor = System.Drawing.Color.Gray;
                txtAppName.Text = "--APPLICATION NAME--";

            }
        }

        private void txtAppName_Enter(object sender, EventArgs e)
        {
            if (txtAppName.Text == "--APPLICATION NAME--")
            {
                txtAppName.Text = "";
            }
            txtAppName.ForeColor = System.Drawing.Color.Black;
        }

        private void notiAHSCT_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            WindowState = FormWindowState.Maximized;
        }

        private void frmconsole_Resize(object sender, EventArgs e)
        {
            //if (FormWindowState.Maximized == WindowState)
              //  Hide();

        }

        #region : LOADING AT APPNAME DROP DOWN WITH THE DESIRED APPLICATION NAMES
        private void rdbBundleB_CheckedChanged(object sender, EventArgs e)
        {
            GlobalData.gATBundle = "B";

            //cmbATAppName.Items.Clear();
            string sCategory = "B";

            if (chkByApp.Checked == true && rdbBundleB.Checked == true)
            {
                DataView dv = GlobalData.dstAplication.Tables[0].DefaultView;
                dv.RowFilter = "AppGroup='" + sCategory + "'";

                cmbATAppName.DataSource = dv;
                cmbATAppName.DisplayMember = "AppName";
                cmbATAppName.ValueMember = "AppName";
            }
            //GlobalData.gCategory = cmbCategory.Text;

        }

        private void rdbBundleD_CheckedChanged(object sender, EventArgs e)
        {
            GlobalData.gATBundle = "D";

            //cmbATAppName.Items.Clear();
            string sCategory = "D";
            if (chkByApp.Checked == true && rdbBundleD.Checked == true)
            {
                DataView dv = GlobalData.dstAplication.Tables[0].DefaultView;
                dv.RowFilter = "AppGroup='" + sCategory + "'";
                cmbATAppName.DataSource = dv;
                cmbATAppName.DisplayMember = "AppName";
                cmbATAppName.ValueMember = "AppName";
            }

        }


        private void rdbBoth_CheckedChanged(object sender, EventArgs e)
        {
            GlobalData.gATBundle = "*";

            if (chkByApp.Checked == true && rdbBoth.Checked == true)
            {
                //cmbATAppName.Items.Clear();
                DataView dv = GlobalData.dstAplication.Tables[0].DefaultView;
                dv.RowFilter = "AppGroup IN ('B','D')";
                cmbATAppName.DataSource = dv;
                cmbATAppName.DisplayMember = "AppName";
                cmbATAppName.ValueMember = "AppName";
            }
        }

        #endregion


        #region : LOAD THE DESIRED AT DATA IN THE GRID
        private void btnGetATData_Click(object sender, EventArgs e)
        {
            try
            {
                btnExportData.Visible = true;
                string appname = cmbATAppName.Text;
                int iRowsReturned;

                Availdt.Rows.Clear();
                #region : BUILDING THE FILTERED QUERY FOR PULLING THE AT DATA

                if (chkByRegion.Checked == true)
                {
                    if (chkByBundle.Checked == true)
                    {
                        if (chkByApp.Checked == true)
                        {
                            //DEEPEST POSSIBLE FILTER APPLICABLE
                            appname = cmbATAppName.Text;
                            GlobalData.gATQuery = "SELECT [ATDate] AS AT_DATE,[TicketReference] AS INCIDENT_NO,[ApplicationName] AS APPLICATION_NAME,[ApplicationCategory] AS BUNDLE,[ServiceClass] AS SERVICE_CLASS,[Country] AS REGION,[UAStartDate] AS UA_START_DATE,[UAEndDate] AS UA_END_DATE,[OutageDuration] AS OUTAGE_DURATION,[Planned] AS PLANNED_ACTIVITY,[DueToAM] AS DUE_TO_AM,[Details] AS DETAILS,[CommentsActions] AS COMMENTS,[ActionOwner] AS OWNER,[UpdatedBy] AS MODIFIER FROM AVAILABILITY WHERE ((Country = '" + GlobalData.gATRegion + "') AND (ApplicationCategory = '" + GlobalData.gATBundle + "') AND (ApplicationName  ='" + appname + "' ))";
                        }
                        else
                        {
                            //NO APPLICATION SELECTED
                            GlobalData.gATQuery = "SELECT [ATDate] AS AT_DATE,[TicketReference] AS INCIDENT_NO,[ApplicationName] AS APPLICATION_NAME,[ApplicationCategory] AS BUNDLE,[ServiceClass] AS SERVICE_CLASS,[Country] AS REGION,[UAStartDate] AS UA_START_DATE,[UAEndDate] AS UA_END_DATE,[OutageDuration] AS OUTAGE_DURATION,[Planned] AS PLANNED_ACTIVITY,[DueToAM] AS DUE_TO_AM,[Details] AS DETAILS,[CommentsActions] AS COMMENTS,[ActionOwner] AS OWNER,[UpdatedBy] AS MODIFIER FROM AVAILABILITY WHERE ((Country = '" + GlobalData.gATRegion + "') AND (ApplicationCategory = '" + GlobalData.gATBundle + "'))";
                        }
                    }

                    else
                    {
                        //NO APPLICATION AND BUNDLE SELECTED 
                        GlobalData.gATQuery = "SELECT [ATDate] AS AT_DATE,[TicketReference] AS INCIDENT_NO,[ApplicationName] AS APPLICATION_NAME,[ApplicationCategory] AS BUNDLE,[ServiceClass] AS SERVICE_CLASS,[Country] AS REGION,[UAStartDate] AS UA_START_DATE,[UAEndDate] AS UA_END_DATE,[OutageDuration] AS OUTAGE_DURATION,[Planned] AS PLANNED_ACTIVITY,[DueToAM] AS DUE_TO_AM,[Details] AS DETAILS,[CommentsActions] AS COMMENTS,[ActionOwner] AS OWNER,[UpdatedBy] AS MODIFIER FROM AVAILABILITY WHERE ((Country = '" + GlobalData.gATRegion + "'))";
                    }

                }

                else
                {
                    if (chkByBundle.Checked == true)
                    {
                        if (chkByApp.Checked == true)
                        {
                            //WITHOUT REGION , ONLY APP AND BUNDLE SELECTED
                            appname = cmbATAppName.Text;
                            GlobalData.gATQuery = "SELECT [ATDate] AS AT_DATE,[TicketReference] AS INCIDENT_NO,[ApplicationName] AS APPLICATION_NAME,[ApplicationCategory] AS BUNDLE,[ServiceClass] AS SERVICE_CLASS,[Country] AS REGION,[UAStartDate] AS UA_START_DATE,[UAEndDate] AS UA_END_DATE,[OutageDuration] AS OUTAGE_DURATION,[Planned] AS PLANNED_ACTIVITY,[DueToAM] AS DUE_TO_AM,[Details] AS DETAILS,[CommentsActions] AS COMMENTS,[ActionOwner] AS OWNER,[UpdatedBy] AS MODIFIER FROM AVAILABILITY WHERE ((ApplicationCategory = '" + GlobalData.gATBundle + "') AND (ApplicationName  ='" + appname + "' ))";
                        }
                        else
                        {
                            //WITHOUT REGION,APPLICATION ONLY BUNDLE TEAM WILL PLAY.
                            GlobalData.gATQuery = "SELECT [ATDate] AS AT_DATE,[TicketReference] AS INCIDENT_NO,[ApplicationName] AS APPLICATION_NAME,[ApplicationCategory] AS BUNDLE,[ServiceClass] AS SERVICE_CLASS,[Country] AS REGION,[UAStartDate] AS UA_START_DATE,[UAEndDate] AS UA_END_DATE,[OutageDuration] AS OUTAGE_DURATION,[Planned] AS PLANNED_ACTIVITY,[DueToAM] AS DUE_TO_AM,[Details] AS DETAILS,[CommentsActions] AS COMMENTS,[ActionOwner] AS OWNER,[UpdatedBy] AS MODIFIER FROM AVAILABILITY WHERE ((ApplicationCategory = '" + GlobalData.gATBundle + "'))";

                        }
                    }
                }
                #endregion


                GlobalData.gApplicatioName = cmbApplication.Text;
                string sCategory;
                sCategory = rdbBundleB.Checked == true ? "B" : "D";
                OleDbCommand cmd = new OleDbCommand();
                //cmd.CommandText = "SELECT [ATDate] AS AT_DATE,[TicketReference] AS INCIDENT_NO,[ApplicationName] AS APPLICATION_NAME,[ApplicationCategory] AS BUNDLE,[ServiceClass] AS SERVICE_CLASS,[UAStartDate] AS UA_START_DATE,[UAEndDate] AS UA_END_DATE,[OutageDuration] AS OUTAGE_DURATION,[Planned] AS PLANNED_ACTIVITY,[DueToAM] AS DUE_TO_AM,[Details] AS DETAILS,[CommentsActions] AS COMMENTS,[ActionOwner] AS OWNER,[UpdatedBy] AS MODIFIER FROM AVAILABILITY WHERE ((ApplicationName = '" + appname + "') AND (ApplicationCategory  ='" + sCategory + "' ))";
                //cmd.CommandText = "SELECT [ATDate] AS AT_DATE,[TicketReference] AS INCIDENT_NO FROM AVAILABILITY WHERE ((ApplicationName = '" + appname + "') AND (ApplicationCategory  ='" + sCategory + "' ))";
                cmd.CommandText = GlobalData.gATQuery;
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.SelectCommand.Connection = GlobalData.GlobalConnection;//conn;
                da.SelectCommand = cmd;
                //DataSet ds1 = new DataSet();
                //DataTable Availdt = new DataTable();
                DataTable dt = new DataTable();
                da.Fill(dt);
                da.Fill(Availdt);

                dgAT.DataSource = Availdt;

                //GETTING THE NO. OF ROWS RETURNED BY THE QUERY
                iRowsReturned = Availdt.Rows.Count;
                if (iRowsReturned == 0)
                {
                    dgAT.Visible = false;
                    grpATResults.Text = "NO RECORDS EXIST FOR THE ABOVE QUERY";
                }
                else
                {
                    dgAT.Visible = true;
                    grpATResults.Text = "DISPLAYING RESULTS : " + iRowsReturned + " rows returned";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Something in the input provided by you is not correct, we suggest you check the input and try again", "Aheem Aheem");
            }

        }

        #endregion

        private void chkByRegion_CheckedChanged(object sender, EventArgs e)
        {
            if (chkByRegion.Checked == true)
            {
                grpATRegions.Enabled = true;
            }

            else
            {
                grpATRegions.Enabled = false;
            }
        }

        private void chkByBundle_CheckedChanged(object sender, EventArgs e)
        {
            if (chkByBundle.Checked == true)
            {
                grpATBundles.Enabled = true;
                chkByApp.Enabled = true;
            }

            else
            {
                grpATBundles.Enabled = false;
                chkByApp.Enabled = false;
            }
        }

        private void chkByApp_CheckedChanged(object sender, EventArgs e)
        {
            if (chkByApp.Checked == true)
            {
                grpATApp.Enabled = true;
            }

            else
            {
                grpATApp.Enabled = false;
            }
        }

        private void rdbUS_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbUS.Checked == true)
            {
                GlobalData.gATRegion = "US";
            }
        }

        private void rdbUK_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbUK.Checked == true)
            {
                GlobalData.gATRegion = "UK";
            }
        }

        private void rdbSWE_CheckedChanged(object sender, EventArgs e)
        {
            if (rdbSWE.Checked == true)
            {
                GlobalData.gATRegion = "SWE";
            }
        }

        private void rdb_H_BundleB_CheckedChanged(object sender, EventArgs e)
        {
            if (rdb_H_BundleB.Checked == true)
            {

                string sCategory;

                sCategory = "B";


                DataView dv = GlobalData.dstAplication.Tables[0].DefaultView;
                dv.RowFilter = "AppGroup='" + sCategory + "'";

                //cmbApplication.Items.Clear();
                cmbApplication.DataSource = dv;
                cmbApplication.DisplayMember = "AppName";
                cmbApplication.ValueMember = "AppName";


                GlobalData.gCategory = "BUNDLE-B";
                GlobalData.gServiceCategory = "B";

                btnFinal.Visible = false;
                btnNext.Visible = false;

            }
        }

        private void rdb_H_BundleD_CheckedChanged(object sender, EventArgs e)
        {

            if (rdb_H_BundleD.Checked == true)
            {
                string sCategory;
                sCategory = "D";

                DataView dv = GlobalData.dstAplication.Tables[0].DefaultView;
                dv.RowFilter = "AppGroup='" + sCategory + "'";

                //cmbApplication.Items.Clear();
                cmbApplication.DataSource = dv;
                cmbApplication.DisplayMember = "AppName";
                cmbApplication.ValueMember = "AppName";


                GlobalData.gCategory = "BUNDLE-D";
                GlobalData.gServiceCategory = "D";

                btnFinal.Visible = false;
                btnNext.Visible = false;

            }
        }

        private void tbSearch_Enter(object sender, EventArgs e)
        {
            #region: POPULATE THE APP NAME AND THE INCIDENT DROPDOWNS WITH APPROPRIATE DATA.

            DataView dv = GlobalData.dstAplication.Tables[0].DefaultView;
            dv.RowFilter = "AppGroup IN ('B','D')";
            cmb_S_AppName.DataSource = dv;
            cmb_S_AppName.DisplayMember = "AppName";
            cmb_S_AppName.ValueMember = "AppName";

            #endregion
        }

        private void btnExportData_Click(object sender, EventArgs e)
        {

            //int DoNothing = 0;
            //int iExportProgress = 0;
            //#region: CHECK IF OUTLOOK IS RUNNING OR NOT(A PRE-REQUISITE)

            //bool IfOutlookNotRunning = true;
            //try
            //{
            //    Process[] ps = Process.GetProcesses();
            //    foreach (Process p in ps)
            //    {
            //        if (p.ProcessName.ToLower().Equals("outlook"))
            //        {
            //            //p.Kill();
            //            //Do Nothing;
            //            IfOutlookNotRunning = false;
            //            break;
            //        }

            //    }
            //    if (IfOutlookNotRunning)
            //    {
            //        MessageBox.Show("You need your Outlook Window Open to send the emails to the clients", "Open your Outlook Now !!!");
            //        //break;
            //    }

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("ERROR " + ex.Message);
            //}
            //#endregion



            //if (IfOutlookNotRunning)
            //{
            //    //DO NOTHING
            //}
            //else
            //{
            //    #region: REQUEST THE USER TO PROVIDE A FILENAME FOR THE EXPORTED DATA
            //    saveFile.Filter = "Excel (*.xls)|*.xls";
            //    if (saveFile.ShowDialog() == DialogResult.OK)
            //    {
            //        if (!saveFile.FileName.Equals(string.Empty))
            //        {
            //            FileInfo f = new FileInfo(saveFile.FileName);
            //            if (f.Extension.Equals(".xls"))
            //            {
            //                GlobalData.gExportFilename = saveFile.FileName;
            //            }
            //            else
            //            {
            //                MessageBox.Show("Invalid file type");
            //            }
            //        }
            //        else
            //        {
            //            MessageBox.Show("You did pick a location to save file to");
            //        }
            //    }
            //    else 
            //    {
            //        DoNothing = 1;
            //    }


            //    MessageBox.Show(GlobalData.gExportFilename);

            //#endregion

            //    Thread.Sleep(2000);
            //    object MissingValue = System.Reflection.Missing.Value;

            //    #region: OPENING THE DESIRED EXCEL FILE AND PREPARING FOR EXPORTING THE INFORMATION
            //    if (DoNothing == 0)
            //    {
            //        #region: INITIALIZE THE LABELS TO PROVIDE THE STATUS OF COMPLETION

            //        MessageBox.Show("KINDLY WAIT TILL THE EXPORT PROCESS GETS COMPLETED", "EXPORTING THE DATA");

            //        lblExportProgress.Text = "Starting Export Process . . .";


            //        #endregion


            //        Excel.Application xl = new Excel.Application();
            //        xl.Visible = false;
            //        //Open the excel sheet
            //        //Excel.Workbook wb = xl.Workbooks.Open(GlobalData.gExportFilename, 0, false, 5, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, true, false, System.Reflection.Missing.Value, false, false, false);

            //        Excel.Workbook wb = xl.Workbooks.Add(MissingValue);
            //        xl.Application.DisplayAlerts = false;
            //        Excel.Sheets wsheets = wb.Sheets;
            //        //Selecting the first sheet
            //        Excel.Worksheet xlwsheet = (Excel.Worksheet)wsheets[1];

            //        int iSheetName;
            //        int iRow;
            //        int iColumn;
            //        string style;

            //    #endregion

            //        #region: CREATING COLUMNS AND ROWS WORKBOOK STYLES
            //        //Creates 2 Custom styles for the workbook These styles are
            //        //  styleColumnHeadings
            //        //  styleRows
            //        //These 2 styles are used when filling the individual Excel cells with the
            //        //DataView values. If the current cell relates to a DataView column heading
            //        //then the style styleColumnHeadings will be used to render the current cell.
            //        //If the current cell relates to a DataView row then the style styleRows will
            //        //be used to render the current cell.


            //        //Style styleColumnHeadings

            //        //////////////////////////////////////////////////////
            //        //try
            //        //{
            //        //    styleColumnHeadings = workbook.Styles["styleColumnHeadings"];
            //        //}
            //        //// Style doesn't exist yet.
            //        //catch
            //        //{
            //        //    styleColumnHeadings = workbook.Styles.Add("styleColumnHeadings", Type.Missing);
            //        //    styleColumnHeadings.Font.Name = "Arial";
            //        //    styleColumnHeadings.Font.Size = 14;
            //        //    styleColumnHeadings.Font.Color = (255 << 16) | (255 << 8) | 255;
            //        //    styleColumnHeadings.Interior.Color = (0 << 16) | (0 << 8) | 0;
            //        //    styleColumnHeadings.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid;
            //        //}

            //        //// Style styleRows
            //        //try
            //        //{

            //        //    styleRows = workbook.Styles["styleRows"];
            //        //}
            //        //// Style doesn't exist yet.
            //        //catch
            //        //{
            //        //    styleRows = workbook.Styles.Add("styleRows", Type.Missing);
            //        //    styleRows.Font.Name = "Arial";
            //        //    styleRows.Font.Size = 10;
            //        //    styleRows.Font.Color = (0 << 16) | (0 << 8) | 0;
            //        //    styleRows.Interior.Color = (192 << 16) | (192 << 8) | 192;
            //        //    styleRows.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid;
            //        //}

            //        #endregion

            //        #region: FILLING DATA INTO EXCEL SHEETS

            //        #region: FILLING THE COLUMN HEADINGS FIRST;
            //        style = "Headings";
            //        for (int i = 1; i < Availdt.Columns.Count; i++)
            //        {
            //            fillExcelCell(xlwsheet, 1, i, Availdt.Columns[i].ToString(),style);
            //        }

            //        #endregion


            //        #region: FILLING THE REMAINING ROWS VALUES HEADINGS FIRST;
            //        style = "Rows";
            //        for (int j = 2; j < Availdt.Rows.Count; j++)
            //        {
            //            for (int i = 1; i < Availdt.Columns.Count; i++)
            //            {
            //                fillExcelCell(xlwsheet,j, i, Availdt.Rows[j][i].ToString() , style);
            //            }

            //            iExportProgress = (100 / Availdt.Rows.Count) * j +2;
            //            lblExportProgress.Text = iExportProgress + "% COMPLETE";
            //        }
            //        #endregion

            //        wb.SaveAs(GlobalData.gExportFilename, Excel.XlFileFormat.xlWorkbookNormal, MissingValue, MissingValue, MissingValue, MissingValue, Excel.XlSaveAsAccessMode.xlExclusive, MissingValue, MissingValue, MissingValue, MissingValue, MissingValue);

            //        xl.Quit();
            //        GC.Collect();

            //        #endregion

            //    }
            //    else
            //    {
            //        //DO NOTHING
            //    }
            //}

            lblExportProgress.Visible = true;
            lblExportProgress.Text = "Starting Export Process . . .";
            

            #region: FROM THE EXPORT  FUNCTION


            int DoNothing = 0;
            #region: CHECK IF OUTLOOK IS RUNNING OR NOT(A PRE-REQUISITE)

            bool IfOutlookNotRunning = true;
            try
            {
                Process[] ps = Process.GetProcesses();
                foreach (Process p in ps)
                {
                    if (p.ProcessName.ToLower().Equals("outlook"))
                    {
                        //p.Kill();
                        //Do Nothing;
                        IfOutlookNotRunning = false;
                        break;
                    }

                }
                if (IfOutlookNotRunning)
                {
                    MessageBox.Show("You need your Outlook Window Open to send the emails to the clients", "Open your Outlook Now !!!");
                    //break;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR " + ex.Message);
            }
            #endregion

            if (IfOutlookNotRunning)
            {
                //DO NOTHING
            }
            else
            {
                #region: REQUEST THE USER TO PROVIDE A FILENAME FOR THE EXPORTED DATA
                saveFile.Filter = "Excel (*.xls)|*.xls";
                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    if (!saveFile.FileName.Equals(string.Empty))
                    {
                        FileInfo f = new FileInfo(saveFile.FileName);
                        if (f.Extension.Equals(".xls"))
                        {
                            GlobalData.gExportFilename = saveFile.FileName;
                        }
                        else
                        {
                            MessageBox.Show("Invalid file type");
                        }
                    }
                    else
                    {
                        MessageBox.Show("You did pick a location to save file to");
                    }
                }
                else
                {
                    DoNothing = 1;
                }


                MessageBox.Show(GlobalData.gExportFilename);

                #endregion

                Thread.Sleep(2000);
                object MissingValue = System.Reflection.Missing.Value;
                pbExportProgress.Visible = true;
                GlobalData.gIfExportWorking = 1;
                btnExportData.Enabled = false;
                lblExportProgress.Text = "Exporting ... ";

                #region: IF CORRECT INFO IS PROVIDED THEN START THE EXPORT PROCESS OR ELSE DONT
                if (DoNothing == 0)
                {
                    BackgroundWorker bgExport = new BackgroundWorker();
                    bgExport.DoWork += new DoWorkEventHandler(bgExport_DoWork);
                    bgExport.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgExport_RunWorkerCompleted);
                    bgExport.RunWorkerAsync();

                }
                else
                {
                    //DO NOTHING

                }
                #endregion

            #endregion


            }

        }

        private void bgExport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //throw new NotImplementedException();
            pbExportProgress.Visible = false;
            GlobalData.gIfExportWorking = 0;
            lblExportProgress.Text = "EXPORT COMPLETED SUCCESSFULLY";
            MessageBox.Show("EXPORT PROCESS COMPLETED SUSSCESSFULLY", "Phewwwwww !!!");
            lblExportProgress.Visible = false;
            pbExportProgress.Visible = false;
            GlobalData.gExportProgressStatus = 0;
                            
        }

        private void bgExport_DoWork(object sender, DoWorkEventArgs e)
        {
            //throw new NotImplementedException();
            ExportAvailability();
        }



        private void spltcntSearch_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void frmconsole_FormClosing(object sender, FormClosingEventArgs e)
        {
            GlobalData.GlobalConnection.Close();
        }

        #region: FILLS THE CELLS INDIVIDUALLY
        public void fillExcelCell(Excel.Worksheet worksheet, int row, int col, string value, string style)
        {
            #region: INSERTING DATA INTO THE EXCEL SHEETS
            Excel.Range rng = (Excel.Range)worksheet.Cells[row, col];
            rng.Select();
            rng.Value2 = value;
            rng.Columns.EntireColumn.AutoFit();
            rng.Borders.Weight = Excel.XlBorderWeight.xlThin;
            rng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            rng.Borders.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
            #endregion

            #region: APPLYING THE APPROPRIATE STYLE TO HEADINGS AND COLUMNS

            if (style == "Headings")
            {

                rng.Font.Name = "Arial";
                rng.Font.Size = 14;
                rng.Font.Color = (255 << 16) | (255 << 8) | 255;
                rng.Interior.Color = (0 << 16) | (0 << 8) | 0;
                rng.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid;
            }
            else
            {
                rng.Font.Name = "Arial";
                rng.Font.Size = 10;
                rng.Font.Color = (0 << 16) | (0 << 8) | 0;
                rng.Interior.Color = (192 << 16) | (192 << 8) | 192;
                rng.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid;
            }


            #endregion



        }
        #endregion


        #region: EXPORTING THE AVAILABILITY INFORAMTION (INDEPENDENT THREAD)

            private void ExportAvailability()
            {

                #region: INITIALIZE THE LABELS TO PROVIDE THE STATUS OF COMPLETION

                //MessageBox.Show("KINDLY WAIT TILL THE EXPORT PROCESS GETS COMPLETED", "EXPORTING AVAILABILITY INFORMATION");

                int iExportProgress = 0;
                
                Excel.Application xl = new Excel.Application();
                xl.Visible = false;
                //Open the excel sheet
                //Excel.Workbook wb = xl.Workbooks.Open(GlobalData.gExportFilename, 0, false, 5, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, true, false, System.Reflection.Missing.Value, false, false, false);
                Object MissingValue = System.Reflection.Missing.Value;
                Excel.Workbook wb = xl.Workbooks.Add(MissingValue);
                xl.Application.DisplayAlerts = false;
                Excel.Sheets wsheets = wb.Sheets;
                //Selecting the first sheet
                Excel.Worksheet xlwsheet = (Excel.Worksheet)wsheets[1];

                int iSheetName;
                int iRow;
                int iColumn;
                string style;

                #endregion

                #region: CREATING COLUMNS AND ROWS WORKBOOK STYLES
                //Creates 2 Custom styles for the workbook These styles are
                //  styleColumnHeadings
                //  styleRows
                //These 2 styles are used when filling the individual Excel cells with the
                //DataView values. If the current cell relates to a DataView column heading
                //then the style styleColumnHeadings will be used to render the current cell.
                //If the current cell relates to a DataView row then the style styleRows will
                //be used to render the current cell.


                //Style styleColumnHeadings

                //////////////////////////////////////////////////////
                //try
                //{
                //    styleColumnHeadings = workbook.Styles["styleColumnHeadings"];
                //}
                //// Style doesn't exist yet.
                //catch
                //{
                //    styleColumnHeadings = workbook.Styles.Add("styleColumnHeadings", Type.Missing);
                //    styleColumnHeadings.Font.Name = "Arial";
                //    styleColumnHeadings.Font.Size = 14;
                //    styleColumnHeadings.Font.Color = (255 << 16) | (255 << 8) | 255;
                //    styleColumnHeadings.Interior.Color = (0 << 16) | (0 << 8) | 0;
                //    styleColumnHeadings.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid;
                //}

                //// Style styleRows
                //try
                //{

                //    styleRows = workbook.Styles["styleRows"];
                //}
                //// Style doesn't exist yet.
                //catch
                //{
                //    styleRows = workbook.Styles.Add("styleRows", Type.Missing);
                //    styleRows.Font.Name = "Arial";
                //    styleRows.Font.Size = 10;
                //    styleRows.Font.Color = (0 << 16) | (0 << 8) | 0;
                //    styleRows.Interior.Color = (192 << 16) | (192 << 8) | 192;
                //    styleRows.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid;
                //}

                #endregion

                #region: FILLING DATA INTO EXCEL SHEETS

                       #region: FILLING THE COLUMN HEADINGS FIRST;
                        style = "Headings";
                        for (int i = 1; i < Availdt.Columns.Count; i++)
                        {
                            fillExcelCell(xlwsheet, 1, i, Availdt.Columns[i].ToString(), style);
                        }

                       #endregion


                    #region: FILLING THE REMAINING ROWS VALUES HEADINGS FIRST;
                        style = "Rows";
                        for (int j = 0; j < Availdt.Rows.Count; j++)
                        {
                            for (int i = 1; i < Availdt.Columns.Count; i++)
                            {
                                fillExcelCell(xlwsheet, j+2, i, Availdt.Rows[j][i].ToString(), style);
                            }

                            iExportProgress = (100 / Availdt.Rows.Count) * j + 2;
                            //lblExportProgress.Text = iExportProgress + "% COMPLETE"; ---> CROSS THREAD HENCE BAD DOUGHNUT FOR YOU
                            GlobalData.gExportProgressStatus = iExportProgress; 
                            
                        }
                    #endregion

                wb.SaveAs(GlobalData.gExportFilename, Excel.XlFileFormat.xlWorkbookNormal, MissingValue, MissingValue, MissingValue, MissingValue, Excel.XlSaveAsAccessMode.xlExclusive, MissingValue, MissingValue, MissingValue, MissingValue, MissingValue);

                xl.Quit();
                
                GC.Collect();

                #endregion

                //}
                //else
                //{
                //DO NOTHING
                //}

            }

        #endregion

            private void btnEnterATInfo_Click(object sender, EventArgs e)
            {
                //SHOWING THE ENTER NEW AVAILABILITY FORM 
                frmNewATInfo frm = new frmNewATInfo();
                frm.Show();

            }

            private void tsbtnFeedback_Click(object sender, EventArgs e)
            {
                Feedback frm = new Feedback();
                frm.Show();
            }

            
    }





}