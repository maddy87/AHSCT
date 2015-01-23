using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Net.Mail;

namespace maddytry1
{
    public class Functions
    {

        public Functions()
        {

        }

        #region : LOAD GLOBAL VALUES WITH THE DETAILS OF THE MOST RECENT UPDATE FOR A GIVEN INCIDENT
        //public void GetUpdateDetails(string TicketRef, ref string TRef, ref string AName, ref string Type, ref string ITitle, ref string BSumm, ref string SPart1, ref string SPart2, ref string Details, ref string BIActual, ref string BIPotential, ref string PActions, ref string DT1, ref string DT2, ref string POC, ref int UNo, ref int Sev)
        public void GetUpdateDetails(string TicketRef)
        {
            try
            {
                //OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OleDB.4.0;Data Source =D:\\DBAS12.mdb");
                //OleDbConnection conn = new OleDbConnection(GlobalData.gAS12_connnectionString);
                //MessageBox.Show(GlobalData.GlobalConnection.State.ToString());
                //conn.Open();

                OleDbCommand cmd = new OleDbCommand("select TicketReference,ApplicationName,Type,UpdateNo,Severity,IncidentTitle,BriefSummary,SummaryPart1,SummaryPart2,Details,BIActual,BIPotential,PreviousActions,DateTime1,DateTime2,POC FROM Master where (TicketReference = '" + TicketRef + "' and UpdateNo = (Select MAX(UpdateNo) from Master where TicketReference = '" + TicketRef + "')) ", GlobalData.GlobalConnection);
                OleDbDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    GlobalData.gCurrTicketRef = dr[0].ToString();
                    GlobalData.gApplicatioName = dr[1].ToString();
                    GlobalData.gCurrent_Type = dr[2].ToString();
                    GlobalData.gUpNo = Convert.ToInt32(dr[3].ToString());
                    GlobalData.gSeverity = Convert.ToInt32(dr[4].ToString());
                    GlobalData.gIncTitle = dr[5].ToString();
                    GlobalData.gBriefSumm = dr[6].ToString();
                    GlobalData.gSummPart1 = dr[7].ToString();
                    GlobalData.gSummPart2 = dr[8].ToString();
                    GlobalData.gDetails = dr[9].ToString();
                    GlobalData.gBIActual = dr[10].ToString();
                    GlobalData.gBIPotential = dr[11].ToString();
                    GlobalData.gPreviousActions = dr[12].ToString();
                    GlobalData.gStartDT = dr[13].ToString();
                    GlobalData.gNextDT = dr[14].ToString();
                    GlobalData.gPOCName = dr[15].ToString();
                }


                dr.Close();
                //conn.Close();
            }
            catch (Exception ex)
            {
                throw ex;
                MessageBox.Show("Problem Retrieving Incident Related Information", "Problemo ");
                Functions fobj = new Functions();
                fobj.ErrorReporting(ex.Message.ToString(), "GetUpdateDetails", "48");
            }
        }

#endregion
        
        #region: INSERT DETAILS INTO MASTER DATABASE (n RECORDS FOR AN INCIDENT)
        public void InsertDetails(string TicketReference, string ApplicationName, string Type, int UpdateNo, int Severity, string IncidentTitle, string BriefSummary, string SummaryPart1, string SummaryPart2, string Details, string BIActual, string BIPotential, string PreviousActions, string DateTime1, string DateTime2, string POC)
        {
            //MessageBox.Show(GlobalData.gNextDT);
                        
            //OleDbConnection conn = new OleDbConnection(GlobalData.gAS12_connnectionString);
            OleDbCommand cmd = new OleDbCommand();
            
            ///Finally implementing DateTime .
            CultureInfo culture = new CultureInfo("en-US");
            DateTime dt1 = Convert.ToDateTime(DateTime1, culture);
            DateTime dt2 = Convert.ToDateTime(DateTime2, culture);

            int id = 0;
            //cmd.CommandText = "insert into Master values(NULL,'" + TicketReference + "')";//,'" + ApplicationName + "','" + Type + "',"+ UpdateNo +"," + Severity + ",'" + IncidentTitle + "','" + BriefSummary + "','" + SummaryPart1 + "','" + SummaryPart2 + "','" + Details + "','" + BIActual + "','" + BIPotential + "','" + PreviousActions + "','" + DateTime1 + "','" + DateTime2 + "','" + POC +"')";
            cmd.CommandText = "insert into Master(TicketReference,ApplicationName,Type,UpdateNo,Severity,IncidentTitle,BriefSummary,SummaryPart1,SummaryPart2,Details,BIActual,BIPotential,PreviousActions,DateTime1,DateTime2,POC) values('" + TicketReference + "','" + ApplicationName + "','" + Type + "'," + UpdateNo + "," + Severity + ",'" + IncidentTitle + "','" + BriefSummary + "','" + SummaryPart1 + "','" + SummaryPart2 + "','" + Details + "','" + BIActual + "','" + BIPotential + "','" + PreviousActions + "','" + dt1 + "','" + dt2 + "','" + POC + "')";
            cmd.Connection = GlobalData.GlobalConnection;
            //conn.Open();
            try
            {
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
                MessageBox.Show("Problem Retrieving Inserting Incident Related Information", "Problemo ");
                Functions fobj = new Functions();
                fobj.ErrorReporting(ex.Message.ToString(), "InsertDetailsMaster", "70");
            }
            finally
            {
                //conn.Close();

            }

        }

        #endregion

        #region: INSERT DETAILS INTO CONSOLE DATABASE ( 1 RECORD FOR AN INCIDENT)
        public void InsertCurrentDetails(string TicketReference, string ApplicationName, string Type, int UpdateNo, int Severity, string IncidentTitle, string BriefSummary, string SummaryPart1, string SummaryPart2, string Details, string BIActual, string BIPotential, string PreviousActions, string DateTime1, string DateTime2, string POC)
        {
            //SqlCommand cmd = new SqlCommand();
            //OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OleDB.4.0;Data Source =D:\\DBAS12.mdb");
            //OleDbConnection conn = new OleDbConnection(GlobalData.gAS12_connnectionString);
            OleDbCommand cmd = new OleDbCommand();

            ///Finally implementing DateTime .
            CultureInfo culture = new CultureInfo("en-US");
            DateTime dt1 = Convert.ToDateTime(DateTime1, culture);
            DateTime dt2 = Convert.ToDateTime(DateTime2, culture);

            //cmd.CommandText = "insert into Console values(,'" + TicketReference + "','" + ApplicationName + "','" + Type + "','" + UpdateNo + "','" + Severity + "','" + IncidentTitle + "','" + BriefSummary + "','" + SummaryPart1 + "','" + SummaryPart2 + "','" + Details + "','" + BIActual + "','" + BIPotential + "','" + PreviousActions + "','" + DateTime1 + "','" + DateTime2 + "','" + POC + "')";
            cmd.CommandText = "insert into Console(TicketReference,ApplicationName,Type,UpdateNo,Severity,IncidentTitle,BriefSummary,SummaryPart1,SummaryPart2,Details,BIActual,BIPotential,PreviousActions,DateTime1,DateTime2,POC) values('" + TicketReference + "','" + ApplicationName + "','" + Type + "'," + UpdateNo + "," + Severity + ",'" + IncidentTitle + "','" + BriefSummary + "','" + SummaryPart1 + "','" + SummaryPart2 + "','" + Details + "','" + BIActual + "','" + BIPotential + "','" + PreviousActions + "','" + dt1 + "','" + dt2 + "','" + POC + "')";
            cmd.Connection = GlobalData.GlobalConnection;
            //conn.Open();
            try
            {
                cmd.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
               // throw ex;
                MessageBox.Show("Problem Retrieving Updating Incident Related Information", "Problemo ");
                Functions fobj = new Functions();
                fobj.ErrorReporting(ex.Message.ToString(), "InsertcurrentDetails", "110");
            }
            finally
            {
              //  conn.Close();
            }

        }

        #endregion

        #region: UPDATE THE DETAILS IN CONSOLE TO THE LATEST UPDATE OF THE INCIDENT 
        public void UpdateCurrentDetails(string TicketReference, string Type, int UpdateNo, int Severity, string UpdateDT)
        {

            //SqlCommand cmd = new SqlCommand();
            //OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OleDB.4.0;Data Source =D:\\DBAS12.mdb");
            //OleDbConnection conn = new OleDbConnection(GlobalData.gAS12_connnectionString);
            OleDbCommand cmd = new OleDbCommand();

            ///Finally implementing DateTime .
            CultureInfo culture = new CultureInfo("en-US");
            DateTime Updatedt1 = Convert.ToDateTime(UpdateDT, culture);

            cmd.CommandText = "UPDATE Console SET Type  = '" + Type + "',UpdateNo = '" + UpdateNo + "',Severity = '" + Severity + "',DateTime2 = '" + Updatedt1 + "'  WHERE TicketReference = '" + TicketReference + "'";
            cmd.Connection = GlobalData.GlobalConnection;
            //conn.Open();
            try
            {
                cmd.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                //throw ex;
                MessageBox.Show("Problem Retrieving Updating Current Incident Related Information", "Problemo ");
                Functions fobj = new Functions();
                fobj.ErrorReporting(ex.Message.ToString(), "UpdateCurrentDetails", "148");
            }
            finally
            {
                //conn.Close();
            }

        }

        #endregion

        #region : UPATING THE AVAILAIBILITY TRACKER ACCORDING TO THE DETAILS IN THE FINAL COMMUNICATION.
        public void UpdateAvailabilityTracker(string UA_StartDate, string UA_EndDate, string OutageTime, string PlannedActivity, string DueTo, string ActionOwner)
        {


            #region: UPDATING THE AVAILABILITY TRACKER USING EXCEL UPDATION METHOD (NOT USED CURRENTLY)
            //Excel.Application xlapp = new Excel.Application();
            //xlapp.Visible = true;
            
            //Excel.Workbook xlwrk = xlapp.Workbooks.Open(strFilename, 0, false, 5, System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, true, false, System.Reflection.Missing.Value, false, false, false);

            //Excel.Sheets xlSheets = xlwrk.Sheets;
            //Excel.Worksheet xlWrkSheet = (Excel.Worksheet)xlSheets[1];

            ////FINDING THE LAST AVAILABLE COLOUMN TO INSERT THE RECORDS 

            //int iLocate = 1;
            //string sCell = "A" + iLocate.ToString();
            //string sTest = xlWrkSheet.get_Range(sCell, sCell).Value2.ToString();


            //try
            //{
            //    while (sCell != "")
            //    {
            //        iLocate++;
            //        sCell = "A" + iLocate.ToString();
            //        sTest = xlWrkSheet.get_Range(sCell, sCell).Value2.ToString();
            //    }
            //}
            //catch (Exception ex)
            //{
            //    //Do Nothing
            //}


            ////INSERTING THE UPDATED VALUES INTO THE AVAILABILITY TRACKER 

            //xlWrkSheet.Cells[iLocate, 1] = DateTime.Now.ToShortDateString();
            //xlWrkSheet.Cells[iLocate, 2] = GlobalData.gCurrTicketRef;
            //xlWrkSheet.Cells[iLocate, 3] = GlobalData.gApplicatioName;
            //xlWrkSheet.Cells[iLocate, 4] = GlobalData.gCategory;
            //xlWrkSheet.Cells[iLocate, 5] = "Service Class";
            //xlWrkSheet.Cells[iLocate, 6] = "Severity" + GlobalData.gSeverity.ToString();
            //xlWrkSheet.Cells[iLocate, 7] = UA_StartDate;
            //xlWrkSheet.Cells[iLocate, 8] = UA_EndDate;
            //xlWrkSheet.Cells[iLocate, 9] = OutageTime;
            //xlWrkSheet.Cells[iLocate, 10] = PlannedActivity;
            //xlWrkSheet.Cells[iLocate, 11] = DueTo;
            //xlWrkSheet.Cells[iLocate, 12] = GlobalData.gIncTitle;
            //xlWrkSheet.Cells[iLocate, 13] = GlobalData.gDetails;
            //xlWrkSheet.Cells[iLocate, 14] = ActionOwner;
            //xlWrkSheet.Cells[iLocate, 15] = GlobalData.gPOCName;


            #endregion

            #region: INSERTING DATA IN THE AVALIBILTY TRACKER  IN THE DATABASE 

            DateTime CurrentAT = new DateTime();
            CurrentAT = DateTime.Now;

            //TEMP
            //GlobalData.gServiceClass = "SERVICE CLASS";

            //OleDbConnection conn = new OleDbConnection(GlobalData.gAS12_connnectionString);
            OleDbCommand cmdAT = new OleDbCommand();
            
            //SETTING UP THE CORRECT FORMAT FOR DATE-TIME
            CultureInfo culture = new CultureInfo("en-US");
            DateTime dt1 = Convert.ToDateTime(UA_StartDate, culture);
            DateTime dt2 = Convert.ToDateTime(UA_EndDate, culture);

            //INSERTING THE DETAILS 
    //        cmd.CommandText = "insert into Console(TicketReference,ApplicationName,Type,UpdateNo,Severity,IncidentTitle,BriefSummary,SummaryPart1,SummaryPart2,Details,BIActual,BIPotential,PreviousActions,DateTime1,DateTime2,POC) values('" + TicketReference + "','" + ApplicationName + "','" + Type + "'," + UpdateNo + "," + Severity + ",'" + IncidentTitle + "','" + BriefSummary + "','" + SummaryPart1 + "','" + SummaryPart2 + "','" + Details + "','" + BIActual + "','" + BIPotential + "','" + PreviousActions + "','" + dt1 + "','" + dt2 + "','" + POC + "')";
    //        cmdAT.CommandText = "insert into AT(ATDate,TicketReference,ApplicationName,ApplicationCategory,ServiceClass,Severity,UAStartDate,UAEndDate,OutageDuration,Planned,DueToAM,Details,CommentsActions,ActionOwner,UpdatedBy,TypeOfIssue,Remarks) values('" + CurrentAT.ToShortDateString()+ "','"+GlobalData.gCurrTicketRef+"','"+GlobalData.gApplicatioName+"','"+GlobalData.gServiceClass+"','"+GlobalData.gServiceClass+"','"+GlobalData.gSeverity.ToString()+"','"+UA_StartDate+"','"+UA_EndDate+"','"+OutageTime+"','"+PlannedActivity+"','"+DueTo+"','"+GlobalData.gIncTitle+"','"+GlobalData.gDetails+"','"+ActionOwner+"','"+GlobalData.gPOCName+"')";
            cmdAT.CommandText = "insert into AVAILABILITY(ATDate,TicketReference,ApplicationName,ApplicationCategory,ServiceClass,Severity,UAStartDate,UAEndDate,OutageDuration,Planned,DueToAM,Details,CommentsActions,ActionOwner,UpdatedBy,Country) values('"+CurrentAT+ "','" + GlobalData.gCurrTicketRef + "','" + GlobalData.gApplicatioName + "','" + GlobalData.gServiceCategory + "','" + GlobalData.gServiceClass + "','" + GlobalData.gSeverity.ToString() + "','" + UA_StartDate + "','" + UA_EndDate + "','" + OutageTime + "','" + PlannedActivity + "','" + DueTo + "','" + GlobalData.gIncTitle + "','" + GlobalData.gDetails + "','" + ActionOwner + "','" + GlobalData.gPOCName + "','"+ GlobalData.gCountryOfSupport + "')";
            cmdAT.Connection = GlobalData.GlobalConnection;

            try
            {
                cmdAT.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
              //  throw ex;
                MessageBox.Show("Problem Inserting Data into Availability Tracker.Please check the inputs provided by you", "Problemo ");
                Functions fobj = new Functions();
                fobj.ErrorReporting(ex.Message.ToString(), "UpdateAvailablity", "241");
            }
            finally
            {
                //conn.Close();
                MessageBox.Show("DETAILS UPADTED IN AVAILABILITY TRACKER");
            }

            #endregion

        }

        #endregion

        #region: ADDING SEPARATE  INTO THE AVAILABILITY TRACKER


        public void NewInfoForAvailabilityTracker(string sATDATE, string sTicketReference, string sAppName, string sAppCategory, string sServiceClass, string sSeverity, string UA_StartDate, string UA_EndDate, string sOutageTime, string sPlannedActivity, string sDueTo, string sDetails, string sCommentsActions, string sActionOwner, string sUpdatedBy, string sTypeOfIssue, string sRemarks, string sCountry)
        {


            #region: INSERTING DATA IN THE AVAILABILTY TRACKER IN THE DATABASE

            DateTime CurrentAT = new DateTime();
            CurrentAT = DateTime.Now;

            //OleDbConnection conn = new OleDbConnection(GlobalData.gAS12_connnectionString);
            OleDbCommand cmdAT = new OleDbCommand();

            //SETTING UP THE CORRECT FORMAT FOR DATE-TIME
            CultureInfo culture = new CultureInfo("en-US");
            DateTime dt1 = Convert.ToDateTime(UA_StartDate, culture);
            DateTime dt2 = Convert.ToDateTime(UA_EndDate, culture);

            //INSERTING THE DETAILS 

            cmdAT.CommandText = "insert into AVAILABILITY (ATDate,TicketReference,ApplicationName,ApplicationCategory,ServiceClass,Severity,UAStartDate,UAEndDate,OutageDuration,Planned,DueToAM,Details,CommentsActions,ActionOwner,UpdatedBy,TypeOfIssue,Remarks,Country) values('" + sATDATE + "','" + sTicketReference + "','" + sAppName + "','" + sAppCategory + "','" + sServiceClass + "','" + sSeverity + "','" + UA_StartDate + "','" + UA_EndDate + "','" + sOutageTime + "','" + sPlannedActivity + "','" + sDueTo + "','" + sDetails + "','" + sCommentsActions + "','" + sActionOwner + "','" + sUpdatedBy + "','" + sTypeOfIssue + "','" + sRemarks + "','" + sCountry + "')";
            cmdAT.Connection = GlobalData.GlobalConnection;


            try
            {
                cmdAT.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                //throw ex;
                MessageBox.Show("Problem Inserting Availability Related Information", "Problemo ");
                Functions fobj = new Functions();
                fobj.ErrorReporting(ex.Message.ToString(), "NewInfoAvailability", "294");
            }
            finally
            {
                //conn.Close();
                MessageBox.Show("DETAILS UPADTED IN AVAILABILITY TRACKER");
            }

            #endregion

        }

        #endregion

        #region: ERROR REPORTING

        //I NEED TO KNOW IF USERS ARE ACTUALLY FACING ANY EXCEPTIONS WHILE USING THE APPPLICATION

        public void ErrorReporting(string error, string functionname, string lineno)
        {
            try
            {
                string sysname = System.Environment.MachineName.ToString();
                string uid1 = System.Environment.UserName.ToString();

                string message  = " --- Error Occurred ---- " + Environment.NewLine + " System/User : "+sysname+ "-"+uid1+ Environment.NewLine+ " Error : "+ error + Environment.NewLine +" InFuntion-LineNo : " +functionname+"-"+lineno  + Environment.NewLine+ " Occured At : "+ DateTime.Now.ToString("dd MMMM yyyy HH:mm ss");
                
                MailMessage Send_Info = new MailMessage();
                Send_Info.From = new MailAddress(uid1 + "@Ikran.com");
                Send_Info.To.Add("rajesh.shetty@Ikran.com");
                Send_Info.Subject = "AHSCT:NOTIFY I am facing an error";
                Send_Info.Body = message;
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
    }
}
