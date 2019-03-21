using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.Odbc;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace EmailPaymentAdvice
{
    public partial class SettingsForm : Form
    {
        private bool isIDE = (Debugger.IsAttached == true);
        private string sendTo;
        private string sendParentTo;
        private string bccTo;
        private string tempCCs = "";
        private string errorEmails = "";
        private string from_email;
        private string from_name;
        private int fromDaysAgo = 0;
        private int toDaysAgo = 0;
        private XmlDocument doc;
        private string originalXml;
        private string settingsFile;

        public SettingsForm()
        {
            InitializeComponent();
        }

        #region properties

        private string _status;
        public string Status
        {
            get { return _status; }
            set { _status = value; ShowStatus(value); }
        }

        #endregion

        #region events

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            // get settings
            try
            {
                string settingsPath = "";
                if(isIDE)
                    settingsPath = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Application.ExecutablePath)));
                else
                    settingsPath = Application.CommonAppDataPath.Remove(Application.CommonAppDataPath.LastIndexOf("."));

                settingsFile = Path.Combine(settingsPath, "Settings.xml");
                doc = new ConfigXmlDocument();
                doc.Load(settingsFile);
                originalXml = doc.OuterXml;

                sendTo = GetSetting(doc, "SendTo");
                sendParentTo = GetSetting(doc, "SendParentTo");
                tempCCs = GetSetting(doc, "SendCC");
                bccTo = GetSetting(doc, "SendBCC");
                errorEmails = GetSetting(doc, "SendErrors");
                var tempFromDaysAgo = GetSetting(doc, "FromDaysAgo");
                var tempToDaysAgo = GetSetting(doc, "ToDaysAgo");
                from_email = GetSetting(doc, "FromEmail");
                from_name = GetSetting(doc, "FromName");
                if(!int.TryParse(tempFromDaysAgo, out fromDaysAgo))
                    fromDaysAgo = 7;
                if(!int.TryParse(tempToDaysAgo, out toDaysAgo))
                    toDaysAgo = 14;
            }
            catch(Exception ex)
            {
                Status = $"Error getting application settings: {ex.Message}";
            }


            // load the schools list
            DataSet ds = GetAllSchools();
            List<string> allSchools = new List<string>();
            foreach(DataRow row in ds.Tables[0].Rows)
            {
                allSchools.Add(row.ItemArray[0].ToString());
            }
            clbSchools.DataSource = allSchools;

            CheckSelectedSchools();

            //this.txtSchools.Text = tempSchools;
            this.txtFromDate.Text = fromDaysAgo.ToString();
            this.txtToDate.Text = toDaysAgo.ToString();
            this.txtToEmail.Text = sendTo;
            this.txtToParentEmail.Text = sendParentTo;
            this.txtFromEmail.Text = from_email;
            this.txtFromName.Text = from_name;
            this.txtCCs.Text = tempCCs;
            this.txtBCCs.Text = bccTo;
            this.txtErrorEmails.Text = errorEmails;
        }

        private void SettingsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(!doc.OuterXml.Equals(originalXml))
            {
                DialogResult result = MessageBox.Show("Changes have been made. Save changes?", "Save Changes?", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question, MessageBoxDefaultButton.Button3);
                if(result == DialogResult.Cancel)
                {
                    e.Cancel = true;
                }
                if(result == DialogResult.Yes)
                {
                    doc.Save(settingsFile);
                    doc = null;
                }
            }
        }

        private void clbSchools_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // collect temp list of items checked
            List<string> checkedItems = new List<string>();
            foreach(var item in clbSchools.CheckedItems)
                checkedItems.Add(item.ToString());

            // add to or remove from temp list
            if(e.NewValue == CheckState.Checked)
                checkedItems.Add(clbSchools.Items[e.Index].ToString());
            else
                checkedItems.Remove(clbSchools.Items[e.Index].ToString());

            // save checked items to xml doc
            string selectedSchools = string.Join(",", checkedItems);
            bool isOk = SetSetting(doc, "Schools", selectedSchools);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            doc.Save(settingsFile);
            originalXml = doc.OuterXml;
            this.Close();
        }

        private void txtFromDate_TextChanged(object sender, EventArgs e)
        {
            bool isOk = SetSetting(doc, "FromDaysAgo", txtFromDate.Text);
        }

        private void txtToDate_TextChanged(object sender, EventArgs e)
        {
            bool isOk = SetSetting(doc, "ToDaysAgo", txtToDate.Text);
        }

        private void txtToEmail_TextChanged(object sender, EventArgs e)
        {
            bool isOk = SetSetting(doc, "SendTo", txtToEmail.Text);
        }

        private void txtToParentEmail_TextChanged(object sender, EventArgs e)
        {
            bool isOk = SetSetting(doc, "SendParentTo", txtToParentEmail.Text);
        }

        private void txtFromEmail_TextChanged(object sender, EventArgs e)
        {
            bool isOk = SetSetting(doc, "FromEmail", txtFromEmail.Text);
        }

        private void txtFromName_TextChanged(object sender, EventArgs e)
        {
            bool isOk = SetSetting(doc, "FromName", txtFromName.Text);
        }

        private void txtCCs_TextChanged(object sender, EventArgs e)
        {
            bool isOk = SetSetting(doc, "SendCC", txtCCs.Text);
        }

        private void txtBCCs_TextChanged(object sender, EventArgs e)
        {
            bool isOk = SetSetting(doc, "SendBCC", txtBCCs.Text);
        }

        private void txtErrorEmails_TextChanged(object sender, EventArgs e)
        {
            bool isOk = SetSetting(doc, "SendErrors", txtErrorEmails.Text);
        }

        #endregion

        #region methods

        delegate void StringArgReturningVoidDelegate(string status);
        private void ShowStatus(string status)
        {
            // InvokeRequired required compares the thread ID of the  
            // calling thread to the thread ID of the creating thread.  
            // If these threads are different, it returns true.  
            if(txtStatus.InvokeRequired)
            {
                StringArgReturningVoidDelegate d = new StringArgReturningVoidDelegate(ShowStatus);
                this.Invoke(d, new object[] { status });
            }
            else
            {
                txtStatus.Text = status;
            }
        }

        private DataSet GetAllSchools()
        {
            DataSet ds = new DataSet();
            try
            {
                using(ODBCClass o = new ODBCClass("MySql"))
                {
                    var procName = "select SchoolID from SchoolInfo";

                    using(OdbcCommand oCommand = o.GetCommand(procName))
                    {
                        using(OdbcDataAdapter oAdapter = new OdbcDataAdapter(oCommand))
                        {
                            oAdapter.Fill(ds);
                        }
                    }
                }
                Status = "";
            }
            catch(Exception ex)
            {
                Status = $"Error getting schools: {ex.Message}";
            }
            return ds;
        }

        private void CheckSelectedSchools()
        {
            // first un-check all
            for(int i = 0; i < clbSchools.Items.Count; i++)
                clbSchools.SetItemChecked(i, false);
            
            // get list of desired schools from settings and check those
            string tempSchools = GetSetting(doc, "Schools");
            List<string> selectedSchools = tempSchools.Split(',').ToList<string>();
            for(int i = 0; i < selectedSchools.Count; i++)
                if(selectedSchools[i] != "")
                    clbSchools.SetItemChecked(clbSchools.Items.IndexOf(selectedSchools[i]), true);
        }

        private string GetSetting(XmlDocument doc, string settingName)
        {
            string response = "";
            try
            {
                response = ((XmlElement)doc.SelectSingleNode($"/Settings/setting[@name='{settingName}']")).GetAttribute("value");
                Status = "";

            }
            catch(Exception ex)
            {
                Status = ex.Message;
            }
            return response;
        }

        private bool SetSetting(XmlDocument doc, string settingName, string newValue)
        {
            bool response = false;
            try
            {
                ((XmlElement)doc.SelectSingleNode($"/Settings/setting[@name='{settingName}']")).SetAttribute("value", newValue);
            }
            catch(Exception ex)
            {
                Status = ex.Message;
            }
            return response;
        }

        #endregion

    }
}
