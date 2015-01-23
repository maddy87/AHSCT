using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;

namespace maddytry1
{
    public partial class frmNewATInfo : Form
    {
        public int gValidate = 0;
        public int gValidateDate = 0;

        public frmNewATInfo()
        {
            InitializeComponent();
        }

        private void rdbBundleD_CheckedChanged(object sender, EventArgs e)
        {
        }


        public void ValidateData()
        {
            //int iReturn = 0;
            gValidate = 0;

            if (rdbSev1.Checked == false && rdbSev2.Checked == false)
            {
                lblRDBErr.Visible = true;
                lblRDBErr.Text = "SEVERITY ???";
                lblRDBErr.ForeColor = System.Drawing.Color.Red;
                gValidate = 1;
            }
            else
            {
                lblRDBErr.Visible = false;
                //GlobalData.gValidate = 0;

            }

            if (rdbBundleB.Checked == false && rdbBundleD.Checked == false)
            {
                MessageBox.Show("Select a Bundle Name", "Huh !!! ");
                gValidate = 1;
            }
            else
            {
                //GlobalData.gValidate = 0;

            }


            if (txtTicketRef.Text == "" || txtTicketRef.Text.Length < 10)
            {
                txtTicketRef.BackColor = System.Drawing.Color.Khaki;
                txtTicketRef.Focus();
                txtTicketRef.BorderStyle = BorderStyle.Fixed3D;
                gValidate = 1;
            }
            else
            {
                txtTicketRef.BackColor = System.Drawing.SystemColors.Window;
                
            }


            if (txtUA_StartDate.Text == "")
            {
                txtUA_StartDate.BackColor = System.Drawing.Color.Khaki;
                txtUA_StartDate.Focus();
                txtUA_StartDate.BorderStyle = BorderStyle.Fixed3D;
                gValidate = 1;
            }
            else
            {
                txtUA_StartDate.BackColor = System.Drawing.SystemColors.Window;
                txtUA_StartDate.Focus();
                //GlobalData.gValidate = 0;
            }

            if (txtUA_EndDate.Text == "")
            {
                txtUA_EndDate.BackColor = System.Drawing.Color.Khaki;
                txtUA_EndDate.Focus();
                txtUA_EndDate.BorderStyle = BorderStyle.Fixed3D;
                
            }
            else
            {
                txtUA_EndDate.BackColor = System.Drawing.SystemColors.Window;
                txtUA_EndDate.Focus();
                //GlobalData.gValidate = 0;
            }

            if (txtActionOwner.Text == "")
            {
                txtActionOwner.BackColor = System.Drawing.Color.Khaki;
                txtActionOwner.Focus();
                txtActionOwner.BorderStyle = BorderStyle.Fixed3D;
                gValidate = 1;
            }
            else
            {
                txtActionOwner.BackColor = System.Drawing.SystemColors.Window;
                txtActionOwner.Focus();
                //GlobalData.gValidate = 0;
            }

            if (txtDetails.Text == "")
            {
                txtDetails.BackColor = System.Drawing.Color.Khaki;
                txtDetails.Focus();
                txtDetails.BorderStyle = BorderStyle.Fixed3D;
                gValidate = 1;
            }
            else
            {
                txtDetails.BackColor = System.Drawing.SystemColors.Window;
                txtDetails.Focus();
                //GlobalData.gValidate = 0;
            }

            if (txtServiceClass.Text == "")
            {
                txtServiceClass.BackColor = System.Drawing.Color.Khaki;
                txtServiceClass.Focus();
                txtServiceClass.BorderStyle = BorderStyle.Fixed3D;
                gValidate = 1;
            }
            else
            {
                txtServiceClass.BackColor = System.Drawing.SystemColors.Window;
                txtServiceClass.Focus();
                //GlobalData.gValidate = 0;
            }

            if (txtOutageDuration.Text == "")
            {
                txtOutageDuration.BackColor = System.Drawing.Color.Khaki;
                txtOutageDuration.Focus();
                txtOutageDuration.BorderStyle = BorderStyle.Fixed3D;
                gValidate = 1;
            }
            else
            {
                txtOutageDuration.BackColor = System.Drawing.SystemColors.Window;
                txtOutageDuration.Focus();
                //GlobalData.gValidate = 0;
            }

            if (txtRemarks.Text == "")
            {
                txtRemarks.BackColor = System.Drawing.Color.Khaki;
                txtRemarks.Focus();
                txtRemarks.BorderStyle = BorderStyle.Fixed3D;
                gValidate = 1;
            }
            else
            {
                txtRemarks.BackColor = System.Drawing.SystemColors.Window;
                txtRemarks.Focus();
                //GlobalData.gValidate = 0;
            }

            if (txtCountry.Text == "")
            {
                txtCountry.BackColor = System.Drawing.Color.Khaki;
                txtCountry.Focus();
                txtCountry.BorderStyle = BorderStyle.Fixed3D;
                gValidate = 1;
            }
            else
            {
                txtCountry.BackColor = System.Drawing.SystemColors.Window;
                txtCountry.Focus();
                //GlobalData.gValidate = 0;
            }

            if (txtTypeOfIssue.Text == "")
            {
                txtTypeOfIssue.BackColor = System.Drawing.Color.Khaki;
                txtTypeOfIssue.Focus();
                txtTypeOfIssue.BorderStyle = BorderStyle.Fixed3D;
                gValidate = 1;
            }
            else
            {
                txtTypeOfIssue.BackColor = System.Drawing.SystemColors.Window;
                txtTypeOfIssue.Focus();
                //GlobalData.gValidate = 0;
            }

            if (txtUpdatedBy.Text == "")
            {
                txtUpdatedBy.BackColor = System.Drawing.Color.Khaki;
                txtUpdatedBy.Focus();
                txtUpdatedBy.BorderStyle = BorderStyle.Fixed3D;
                gValidate = 1;
            }
            else
            {
                txtUpdatedBy.BackColor = System.Drawing.SystemColors.Window;
                txtUpdatedBy.Focus();
                //GlobalData.gValidate = 0;
            }

            if (txtCommentActions.Text == "")
            {
                txtCommentActions.BackColor = System.Drawing.Color.Khaki;
                txtCommentActions.Focus();
                txtCommentActions.BorderStyle = BorderStyle.Fixed3D;
                gValidate = 1;
            }
            else
            {
                txtCommentActions.BackColor = System.Drawing.SystemColors.Window;
                txtCommentActions.Focus();
                //GlobalData.gValidate = 0;
            }



            }

            //return iReturn;
        

        public void ValidateDate()
        {
            //strign sTempStartDT = StartTimeDT.to
            //DateTime dt;

            string sUA_StartDT = txtUA_StartDate.Text.Trim();
            string sUA_EndDT = txtUA_EndDate.Text.Trim();
            
            string sError = "";
            try
            {
                DateTime dt = DateTime.Parse(sUA_StartDT);
            }
                catch (Exception ex)
                {
                //MessageBox.Show("Start Date\\Time is not entered in proper format");
                sError = " Start Date\\Time";
                gValidateDate = 1;
            }

            try { DateTime dt = DateTime.Parse(sUA_EndDT); }
            catch (Exception ex)
            {
                //MessageBox.Show("Next Date\\Time is not entered in proper format");
                sError = " Next Date\\Time";
                gValidateDate = 1;
            }
            if (gValidateDate == 1)
            {
                MessageBox.Show(sError + " are not in proper format ");
            }


        }

        private void frmNewATInfo_Load(object sender, EventArgs e)
        {
            lblRDBErr.Visible = false;
            //rdbDueToAMNo.Checked = true;
            //rdbPlannedNo.Checked = true;
            cmbPlanned.SelectedText = "NO";
            cmbDueToAM.SelectedText = "NO";
               
            txtUA_StartDate.Text = DateTime.Now.ToString("dd MMMM yyyy HH:mm"); 
        }

        private void btnUpdateAvailabilityInfo_Click(object sender, EventArgs e)
        {
            #region: VALIDATIONS OF THE PROVIDED INPUTS

            ValidateData();
            ValidateDate();
            if (gValidate == 1 || gValidateDate == 1)
            {
                //DO NOTHING AS INVALID DATA HAS BEEN PROVIDED.
                gValidateDate = 0;
            }

            #endregion


            else
            {
                //UPDATING THE VALID INFORMATION INTO THE DATABASE
                string sCategory = rdbBundleB.Checked == true ? "B" : "D" ;
                string sPlanned = cmbPlanned.Text;
                string sDueToAM = cmbDueToAM.Text;
                string sSev = rdbSev1.Checked == true ? "1" : "2" ;
                
                Functions obj = new Functions();
                try
                {
                    obj.NewInfoForAvailabilityTracker(DateTime.Now.ToString(), txtTicketRef.Text,cmbApplication.Text, sCategory, txtServiceClass.Text, sSev, txtUA_StartDate.Text, txtUA_EndDate.Text, txtOutageDuration.Text, sPlanned, sDueToAM, txtDetails.Text, txtCommentActions.Text, txtActionOwner.Text, txtUpdatedBy.Text, txtTypeOfIssue.Text, txtRemarks.Text, txtCountry.Text);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.Message, "HUH !!!");
                    MessageBox.Show("Error updating Availability Information", "Error :( ");
                }

            }
        }

        private void rdbBundleB_CheckedChanged_1(object sender, EventArgs e)
        {
            try
            {
                if (rdbBundleB.Checked == true)
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
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error configuring the Application DropDown", "Error :( ");
            }

        }

        private void rdbBundleD_CheckedChanged_1(object sender, EventArgs e)
        {

            if (rdbBundleD.Checked == true)
            {
                try
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
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error configuring the Application DropDown", "Error :( ");
                }


            }
        }

        private void txtUA_EndDate_Leave(object sender, EventArgs e)
        {
            try
            {
                CultureInfo cul = new CultureInfo("en-US");
                DateTime dtStart = Convert.ToDateTime(txtUA_StartDate.Text, cul);
                DateTime dtEnd = Convert.ToDateTime(txtUA_EndDate.Text, cul);
                System.TimeSpan ts = dtEnd - dtStart;
                string sOutage = ts.Days.ToString() + " days " + ts.Hours.ToString() + ":" + ts.Minutes.ToString() + ":" + ts.Seconds.ToString();
                txtOutageDuration.Text = sOutage;
                
            }
            catch (Exception ex)
            {
                txtOutageDuration.Text = "<---incorrect dates --->";

            }

        }

        
    }
}
