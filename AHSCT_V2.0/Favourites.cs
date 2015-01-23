using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace maddytry1
{
    public partial class Favourites : Form
    {
        public Favourites()
        {
            InitializeComponent();
        }

        private void Favourites_Load(object sender, EventArgs e)
        {
            btnCreateSimilar.Enabled = false;

            //THE FUNCTION WILL POPULATE THE FAVOURTES GRID 
            string sType = "Notification";
            OleDbConnection conn;
            //string connstr = "Provider=Microsoft.Jet.OleDB.4.0;Data Source =D:\\DBAS12.mdb";
            string connstr = GlobalData.gAS12_connnectionString;
            conn = new OleDbConnection(connstr);
            OleDbCommand cmd = new OleDbCommand();
            
            string a = "Notification";
            cmd.CommandText = "Select DISTINCT [TicketReference] AS INCIDENT_NO,[IncidentTitle] AS INCIDENT_TITLE,[ApplicationName] AS APPLICATION_NAME,[Severity] AS SEVERITY,[DateTime1] AS START_DATE  FROM Master WHERE ApplicationName = '" + GlobalData.gApplicatioName + "' AND Type = '" + a + "'";//where AppGroup = '" + cmbCategory.Text + "'";            //cmd.CommandText = "Select TicketReference,IncidentTitle From Master Where ApplicationName = "+GlobalData.gApplicatioName +"";//AS INCIDENT_TITLE,[ApplicationName] AS APPLICATION_NAME,[Severity] AS SEVERITY,[DateTime1] AS START_DATE  FROM Console WHERE TicketReference IN (SELECT TicketReference from console where ApplicationName = "+GlobalData.gApplicatioName+")"; 
            //cmd.CommandText = "Select DISTINCT(AppName,AppGroup) from Application ";

            OleDbDataAdapter da = new OleDbDataAdapter(cmd);
            da.SelectCommand.Connection = conn;
            da.SelectCommand = cmd;
            DataTable dt = new DataTable();
            da.Fill(dt);

            //CHECKING IF THERE ARE ANY FAVOURITES THAT EXIST FOR THIS PARTICULAR APPLICATION.
            int iCheckRowsReturned = dt.Rows.Count;

            if (iCheckRowsReturned == 0)
            {

                grpFavourites.Visible = false;

                MessageBox.Show(" NO FAVOURITES AVAILAIBLE FOR THIS APPLICATION");

                //Generating a new notifcation

                //btnFinal.Enabled = false;
                //btnNext.Enabled = false;

                GlobalData.gNoti_Type = "N";
                GlobalData.gStatus = "Notification";
                GlobalData.gType = "Notification";
                //GlobalData.gCategory = cmbCategory.Text;
                //GlobalData.gApplicatioName = cmbApplication.Text;
                GlobalData.gUpNo = 0;

                Generate gObj = new Generate();
                gObj.ShowDialog();
                gObj.Dispose();


                this.Close();
                this.Dispose();
            }
            else
            {
                dgFavourites.DataSource = dt;
            }
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
            this.Close();
            this.Dispose();
        }

        private void btnCreateNew_Click(object sender, EventArgs e)
        {
            //Generating a new notifcation

            //btnFinal.Enabled = false;
            //btnNext.Enabled = false;

            GlobalData.gNoti_Type = "N";
            GlobalData.gStatus = "Notification";
            GlobalData.gType = "Notification";
            //GlobalData.gCategory = cmbCategory.Text;
            //GlobalData.gApplicatioName = cmbApplication.Text;
            GlobalData.gUpNo = 0;

            Generate gObj = new Generate();
            gObj.ShowDialog();
            gObj.Dispose();

            this.Close();
            this.Dispose();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
            this.Dispose();
        }

        private void dgFavourites_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            btnCreateSimilar.Enabled = true;
            GlobalData.gCurrTicketRef = dgFavourites.CurrentRow.Cells[0].Value.ToString();
        }
    }
}
