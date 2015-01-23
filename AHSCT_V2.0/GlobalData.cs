using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Globalization;
using System.Windows.Forms;
using System.Threading;
using System.Data;

namespace maddytry1
{
    public static class GlobalData
    {
        //Declaring Global Variables to be used in the application

        public static string gStatus;
        public static string gCategory;
        public static string gApplicatioName;
        public static int gSeverity;
        public static string gType;
        public static string gCurrent_Type;    
        public static string gNoti_Type; // Types : N , U , F , FF , FN 
        public static string gIncTitle;
        public static string gBriefSumm;
        public static string gSummPart1;
        public static string gSummPart2;
        public static string gDetails;
        public static string gBIActual;
        public static string gBIPotential;
        public static string gPreviousActions;
        public static string gStartDT;
        public static string gNextDT;
        public static string gEstimatedDT;
        public static string gPOCName;
        public static string gPOCEmailId;
        public static int gUpNo;
        public static string gCurrTicketRef;
        public static int gCurrentRow;
        public static string gFilename;
        public static string gFileLoaction;
        public static string gSeverity1DL;
        public static string gSeverity2DL;
        //INFY DB LOCATION
        
        public static string gAS12_connnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source =D:\\DBAHSCT.mdb";
        //public static string gAS12_connnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =D:\\DBAHSCT.mdb";
        //public static string gAS12_connnectionString = "D:\\DBAHSCT.mdb";


        
        public static string gDatabaseLocation = "D:\\DBAHSCT.mdb";
        
        //CONNECTIVITY ISSUES
        public static int gUnableToConnect = 0;

        //ALTERNATIVE FOR THE MACROS IN THE WORD FILE.DRAFTING A NEW MAIL WITH COMPLETE CONTENTS OF THE WORD FILE INTO THE BODY.
        public static string gEmailBody;

        //  RESOLVING THE DATE TIME ISSUES ( ONCE AND FOR ALL )
        public static string gTimeZone;
        public static string gTempStartDT;
        public static string gTempNextDT;
        public static string gTempResolvedDT;

        //GLOBALDATA FOR THE GENERATE NOTIFICATION STATUS BAR
        public static string gServiceLevel;
        public static string gCountryOfSupport;
        public static string gClientArea;
        public static string gUserBase;
        public static string OFFL2Manager;
        public static string ONL2Manager;

        //GLOBAL DATA NEEDED FOR THE AVAILABILITY TRACKER
        public static string gPlannedActivity;
        public static string gDueToAM;
        public static string gOutageDuration;
        public static string gActionOwner;
        public static int gUpdateAT;
        public static string gServiceClass;
        public static string gServiceCategory;

        //GLOBAL DATA INCLUDED FOR CURRENT STATUS
        public static int gTempRowCount;

        //GLOBAL DATA INCLUDED FOR DATABASE CONNECTION
        public static OleDbConnection GlobalConnection = new OleDbConnection(GlobalData.gAS12_connnectionString);

        //GLOBAL DATA INCLUDED FOR GLOBAL DATASETS
        public static DataSet dstAplication = new DataSet();
        public static DataSet dstMaster = new DataSet();
        public static DataSet dstConsole = new DataSet();

        //GLOBAL DATA INCLUDED FOR AVAILABILITY TRACKER AUTOMATION.
        public static string gATQuery;
        public static string gATRegion;
        public static string gATBundle;
        public static string gExportFilename;
        public static int gExportProgressStatus;
        public static int gIfExportWorking;

        //GLOBAL VARIABLE FOR SPLASH SCREEN
        public static int gSplashComplete;
        public static string gCurrentFileName;

        //GLOBAL VALUES FOR HOLDING THE VALIDATION INFORMATION

        public static int gValidate;
        public static int gValidateDate;

        //GLOBAL VALUE FOR STORING THE CURRENT USER NAME.
        public static string gCurrentUser;

        //GLOBAL VALUE FOR UPDATING THE CURRENT STATUS OF THE DATA
        public static int gUpdateCurrentInfo;

        //GLOBAL VALUE FOR SEND STATUS OF NOTIFICATIONS
        public static int gHasNotificationSent;
    }

    class Operations
    {

        private DataGridView dg_Ops;

        Thread th_LoadApplicationDataSets;
        Thread th_LoadMasterDataSets;
        Thread th_LoadConsoleDataSets;

        public Operations(DataGridView dg_Local)
        {
            dg_Ops = dg_Local;
        }

        public Operations()
        {

        }

        private delegate void DisplayCurrentStatusDelegate();

       // #region LOADING THE DATA INTO ALL RESPECTIVE DATASETS

        public void CalculateCurrentStatus()
        {
            try
            {
                int rowcount = 0;
                string typ = "Final";
                OleDbConnection conn = new OleDbConnection(GlobalData.gAS12_connnectionString);
                OleDbCommand curr_cmd = new OleDbCommand("select TicketReference,ApplicationName,Type,UpdateNo,Severity,IncidentTitle,DateTime1,DateTime2,POC FROM Console where Type <> '" + typ + "'", conn);
                conn.Open();
                OleDbDataReader drr = curr_cmd.ExecuteReader();

                //COUNTING THE NO. OF ROWS TURNED.
                while (drr.Read())
                { rowcount++; }

                GlobalData.gTempRowCount = rowcount;

                drr.Close();

                OleDbDataReader dr1 = curr_cmd.ExecuteReader();
                //dg_Ops.Width = this.Width - 10;
                dg_Ops.Rows.Add(50);
                int rowc = 0;
                while (dr1.Read())
                {
                    //MessageBox.Show(dr1[0].ToString());
                    for (int i = 0; i < 9; i++)
                    {
                        dg_Ops.Rows[rowc].Cells[i].Value = dr1[i].ToString();

                    }
                    //dg_CurrentInfo.Rows[rowc].Cells[8].Value = GetRemainingTime(dr1[7].ToString());
                    
                    rowc++;

                }
            }
            catch (Exception ex)
            {
               // MessageBox.Show("Error Calculating the Current Status","Problemo");
            }
        }

        public void LoadCurrentStatus()
        {
            try
            {
                dg_Ops.Invoke(new DisplayCurrentStatusDelegate(CalculateCurrentStatus));
            }
            catch (Exception ex)
            {
                if (GlobalData.gUnableToConnect == 1)
                {
                    //MessageBox.Show("Error Inkoving Calculated Current Status", "Problemo");
                }
                else { MessageBox.Show("Error Inkoving Calculated Current Status", "Problemo"); }
            }
        }


        public void LoadApplicationDataset()
        {
           
                //GlobalData.dstAplication.Clear();
                OleDbConnection conn = new OleDbConnection(GlobalData.gAS12_connnectionString);
                OleDbCommand cmd = new OleDbCommand("Select * from Application order by AppName");
                OleDbDataAdapter dadp = new OleDbDataAdapter();
                dadp.SelectCommand = cmd;
                dadp.SelectCommand.Connection = conn;
                try
                {

                    dadp.Fill(GlobalData.dstAplication);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message.ToString(), "Well this is embarrasing");
                  //  MessageBox.Show("Error loading Application Datset", "Another Problemo");
                }
            
            

        }

        public void LoadConsoleDataset()
        {
            try
            {
                OleDbConnection conn = new OleDbConnection(GlobalData.gAS12_connnectionString);
                OleDbCommand cmd = new OleDbCommand("Select * from Console");
                OleDbDataAdapter dadp = new OleDbDataAdapter();
                dadp.SelectCommand = cmd;
                dadp.SelectCommand.Connection = conn;
                dadp.Fill(GlobalData.dstConsole);
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Problem Loading Datafrom Database ", "Another Problemo");
            }

        }

        public void LoadMasterDataset()
        {
            try
            {
                OleDbConnection conn = new OleDbConnection(GlobalData.gAS12_connnectionString);
                OleDbCommand cmd = new OleDbCommand("Select * from Master where Type = 'Notification' ");
                OleDbDataAdapter dadp = new OleDbDataAdapter();
                dadp.SelectCommand = cmd;
                dadp.SelectCommand.Connection = conn;
                dadp.Fill(GlobalData.dstConsole);
            }
            catch (Exception ex)
            {
                //MessageBox.Show("Problem Loading Datafrom Database ", "Another Problemo");

            }


        }
        
        public void LoadAllDatasets()
        {
            try
            {

                th_LoadApplicationDataSets = new Thread(new ThreadStart(LoadApplicationDataset));
                //th_LoadConsoleDataSets = new Thread(new ThreadStart(LoadConsoleDataset));
                //th_LoadMasterDataSets = new Thread(new ThreadStart(LoadMasterDataset));
                th_LoadApplicationDataSets.Start();
                //th_LoadConsoleDataSets.Start();
                //th_LoadMasterDataSets.Start();
            }
            catch (Exception ex)
            {
                //MessageBox.Show("There seems to be and issue with the database. Please try re-starting the application.", "Another Problemo");
            }

        }

        public void ReloadAllDataSets()
        {
            try
            {
                th_LoadApplicationDataSets.Abort();
                //th_LoadConsoleDataSets.Abort();
                //th_LoadMasterDataSets.Abort();
                LoadAllDatasets();
            }
            catch (Exception ex)
            {
               // MessageBox.Show("There seems to be and issue with the database. Please try re-starting the application.", "Another Problemo");
            }

        }

        public void LoadConsole()
        {
            frmconsole frm = new frmconsole();
            frm.Show();
        }

        public void GetApplicationInfo()
        {
            try
            {
                OleDbConnection conn = new OleDbConnection(GlobalData.gAS12_connnectionString);
                OleDbCommand cmd = new OleDbCommand("Select ServiceLevel,CountryOfSupport,ClientArea,UserBase,ONL2Manager,OFL2Manager,Severity1DL,Severity2DL from Application where AppName = '" + GlobalData.gApplicatioName + "'", GlobalData.GlobalConnection);
                OleDbDataReader dr = cmd.ExecuteReader();
                dr.Read();

                GlobalData.gServiceClass = dr[0].ToString();
                GlobalData.gCountryOfSupport = dr[1].ToString();
                GlobalData.gClientArea = dr[2].ToString();
                GlobalData.gUserBase = dr[3].ToString();
                GlobalData.ONL2Manager = dr[4].ToString();
                GlobalData.OFFL2Manager = dr[5].ToString();
                GlobalData.gSeverity1DL = dr[6].ToString();
                GlobalData.gSeverity2DL = dr[7].ToString();
            }

            catch (Exception ex)
            {
                MessageBox.Show("Not able to retrive the Application Information.", "Another Problemo");
            }

        }

    }

}
