using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Net.Mail;
using System.Diagnostics;

namespace maddytry1
{
    public partial class Splash : Form
    {
        string sStatus;
        string sAppend;
        //frmconsole frm = new frmconsole();
        Operations obj = new Operations();
        Thread th_OpenConnection;
        Thread th_LoadingAllDatasets;
        public Splash()
        {
            InitializeComponent();
            
        }

        private void Splash_Load(object sender, EventArgs e)
        {

            #region: GET DETAILS OF CURRENTLY LOGGED IN USER
            //char[] cTrim = new char[]{'I','T','L','INF0SYS','\\'};
            //GlobalData.gCurrentUser = System.Security.Principal.WindowsIdentity.GetCurrent().Name.TrimStart(cTrim);
            //GlobalData.gCurrentUser = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Remove(0, 11);
            GlobalData.gCurrentUser = System.Environment.UserName.ToString();
            GlobalData.gCurrentUser = GlobalData.gCurrentUser.ToUpper();
            //MessageBox.Show(GlobalData.gCurrentUser);

            //frm.Show();
            //frm.WindowState = FormWindowState.Minimized;
            //frm.Visible = false;

            GlobalData.gCurrentFileName = "AHSCT_V134.exe";

              

            #endregion

            #region: CONNECT TO DATABASE ON LOAD

            pbStatsus.Minimum = 0;
                pbStatsus.Maximum = 100;
                pbStatsus.Step = 10;

                timSplash.Interval = 100;
                timSplash.Enabled = true;
                //GlobalData.GlobalConnection.Open();

                
              
                //CREATING A SEPARATE THREAD TO MAKE THE APPLICATION RESPONSIVE
                th_OpenConnection = new Thread(new ThreadStart(OpenDatabaseConnection));

            #endregion

            #region: CREATING SEPARATE THREADS TO POPULATE THE DATASETS

                try
                {

                    //obj.LoadAllDatasets();
                    th_LoadingAllDatasets = new Thread(new ThreadStart(obj.LoadAllDatasets));
                    th_LoadingAllDatasets.Start();
                    pbStatsus.Visible = false;
                    Thread.Sleep(3500);
                    pbStatsus.Visible = true;
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message.ToString(), "Well this is embarrasing");
                    MessageBox.Show("AHSCT has encountered issues while tryng to access the DB. Please try restarting the application.", "Well this is embarrasing");
                    GlobalData.gUnableToConnect = 1;
                }

            #endregion
        }

        private void timSplash_Tick(object sender, EventArgs e)
        {
           
            if (pbStatsus.Value <= 90)
            {
                    switch (pbStatsus.Value)
                    {

                        case 10: lblStatus.Text = "Initializing Modules . . .";
                            break;
                        case 20: lblStatus.Text = "Loading Modules . . .";
                            break;
                        case 30: lblStatus.Text = "Connecting to Database . . .";
                            break;
                        case 40: lblStatus.Text = "Connecting to Database . . .";
                            break;
                    }

                    if (pbStatsus.Value == 40)
                    {
                        try
                        {

                            //MessageBox.Show(GlobalData.GlobalConnection.State.ToString());
                            lblStatus.Text = "Connecting to Database . . .";
                            GlobalData.GlobalConnection.Open();
                            //th_OpenConnection.Start();
                            string sConnectionState = GlobalData.GlobalConnection.State.ToString();
                            while (sConnectionState == "Closed")
                            {
                                if (lblStatus.Text == "Connecting to Database . . .")
                                {
                                    lblStatus.Text = "Connecting to Database .";
                                    Thread.Sleep(1000);
                                }
                                else if (lblStatus.Text == "Connecting to Database .")
                                {
                                    lblStatus.Text = "Connecting to Database . .";
                                    Thread.Sleep(100);
                                }
                                else
                                {
                                    lblStatus.Text = "Connecting to Database . . .";
                                    Thread.Sleep(100);
                                }

                                pbStatsus.Step = 1;
                                pbStatsus.PerformStep();
                                sConnectionState = GlobalData.GlobalConnection.State.ToString();
                            }

                            lblStatus.Text = "Database Connection Successfull.";
                            pbStatsus.Value = 80;
                            timSplash.Interval = 50;
                            //th_OpenConnection.Abort();
                        }
                        catch (Exception ex)
                        {
                            System.IO.FileInfo fi = new System.IO.FileInfo(GlobalData.gDatabaseLocation);
                            if (fi.Exists)
                            {
                                //DO NOTHING
                                timSplash.Enabled = false;
                                MessageBox.Show(ex.Message.ToString(), "Well this is embarrasing");
                                //MessageBox.Show("AHSCT is not able to connect to the database server. Please make sure your are connected to the network or try again later.", "Well this is embarrasing");
                                this.Close();
                                this.Dispose();
                                GlobalData.gUnableToConnect = 1;
                                GlobalData.gSplashComplete = 1;
                            }
                            else
                            {
                                timSplash.Enabled = false;
                                this.Close();
                                this.Dispose();
                                GlobalData.gUnableToConnect = 1;
                                GlobalData.gSplashComplete = 1;
                                MessageBox.Show("AHSCT DATABASE HAS BEEN BROUGHT DOWN DUE TO PLANNED MAITAINENCE.WE WILL BE BACK SOON ", "Oooppps !!!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                Application.Exit();
                            }
                            
                            
                         }
                    }
                    pbStatsus.PerformStep();
           
               
            }
                       

            else
            {
                timSplash.Enabled = false;
                GlobalData.gSplashComplete = 1;
                this.Close();
                                               
            }
        }

        public void OpenDatabaseConnection()
        {
            try
            {
                GlobalData.GlobalConnection.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to Connect to the Database.", "Well this is embarrasing");
            }
        }
    }
}
