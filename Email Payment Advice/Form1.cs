using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SendGrid;
using SendGrid.Helpers.Mail;
using Aimm.Logging;
using System.IO;
using System.Xml;
using System.Diagnostics;
using System.Globalization;
using System.Threading;
using System.Security.Principal;
using System.Security.AccessControl;

namespace EmailPaymentAdvice
{
    public partial class Form1 : Form
    {
        private List<string> CCsTo = new List<string>();
        private List<string> BCCsTo = new List<string>();
        private List<string> ERRsTo = new List<string>();
        private List<string> errorList = new List<string>();
        private string sendTo;
        private string sendParentTo;
        private Response response;
        private string textBody;
        private string htmlBody;
        private string curStudentFirst;
        private string curStudentLast;
        private string curSchoolName;
        private string curSchoolID;
        private string curSchoolContact;
        private string curSchoolContactEmail;
        private string curStudentEmail;
        private string curParentEmail;
        private string curSchoolCcEmail;
        private string curSchoolErrorEmail;
        private int fromDaysAgo = 0;
        private int toDaysAgo = 0;
        private string apiKey;
        DateTime fromDate;
        DateTime toDate;
        private string status = null;
        string from_email;
        string from_name;
        string fromEmail;
        string fromName;
        private string msg = "";
        private bool isIDE = (Debugger.IsAttached == true);
        string[] args = Environment.GetCommandLineArgs();
        private bool isAutoRun;
        private bool isInstall;
        private string settingsFile;
        private XmlDocument doc;
        private int problems = 0;

        #region boilerplate_text
        private string text_body = "Dear {0} {1},\n\nYour Financial Aid to attend {2} has been posted "
                    + "to your student account as follows:\n\n{3}\n\n"
                    + "Please understand that as a student you have the right to cancel all or a portion of a loan "
                    + "or loan disbursements.  The request must be in writing and must be received by the school "
                    + "within 30 days from the date of the loan disbursement.  If you chose this option, you may "
                    + "owe the school money which may be payable at the time of cancellation.\n\n"
                    + "If you have any questions, please contact {4}’s Financial Aid Department.\n\n"
                    + "Please do not reply to this email because this email box is not monitored.";

        private string parent_text_body = "Dear Parent(s) of {0} {1},\n\nYour Financial Aid to attend {2} has been posted "
            + "to your student's account as follows:\n\n{3}\n\n"
            + "Please understand that as a parent you have the right to cancel all or a portion of a PLUS loan "
            + "or loan disbursements.  The request must be in writing and must be received by the school "
            + "within 30 days from the date of the loan disbursement.  If you chose this option, you may "
            + "owe the school money which may be payable at the time of cancellation.\n\n"
            + "If you have any questions, please contact {4}’s Financial Aid Department.\n\n"
            + "Please do not reply to this email because this email box is not monitored.";

        private string html_body = "<p>Dear {0} {1},</p><p>Your Financial Aid to attend {2} has been posted "
                    + "to your student account as follows:</p><p><pre>{3}</pre></p>"
                    + "<p>Please understand that as a student you have the right to cancel all or a portion of a loan "
                    + "or loan disbursements.  The request must be in writing and must be received by the school "
                    + "within 30 days from the date of the loan disbursement.  If you chose this option, you may "
                    + "owe the school money which may be payable at the time of cancellation.</p>"
                    + "<p>If you have any questions, please contact {4}’s Financial Aid Department.</p>"
                    + "<p>Please do not reply to this email because this email box is not monitored.</p>";

        private string parent_html_body = "<p>Dear Parent(s) of {0} {1},</p><p>Your Financial Aid to attend {2} has been posted "
                    + "to your student's account as follows:</p><p><pre>{3}</pre></p>"
                    + "<p>Please understand that as a parent you have the right to cancel all or a portion of a PLUS loan "
                    + "or loan disbursements.  The request must be in writing and must be received by the school "
                    + "within 30 days from the date of the loan disbursement.  If you chose this option, you may "
                    + "owe the school money which may be payable at the time of cancellation.</p>"
                    + "<p>If you have any questions, please contact {4}’s Financial Aid Department.</p>"
                    + "<p>Please do not reply to this email because this email box is not monitored.</p>";

        private string text_error_body = "Dear EmailPaymentAdvice Administrator,\n\nThe following errors have occurred "
                    + "while processing emails for {0}:\n\n{1}\n\n"
                    + "Please do not reply to this email because this email box is not monitored.";

        private string html_error_body = "<p>Dear EmailPaymentAdvice Administrator,</p><p>The following errors have occurred "
                    + "while processing emails for {0}:</p><p><pre>{1}</pre></p>"
                    + "<p>Please do not reply to this email because this email box is not monitored.</p>";
        #endregion

        public Form1()
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

        private void Form1_Load(object sender, EventArgs e)
        {
            LogIt.LogMethod();

            // setup ToolTips
            ToolTip toolTip1 = new ToolTip();
            toolTip1.AutoPopDelay = 5000;
            toolTip1.InitialDelay = 1000;
            toolTip1.ReshowDelay = 500;
            toolTip1.ShowAlways = true;
            toolTip1.SetToolTip(this.lblSettings, "Settings");

            lblVersion.Text = $"v {Application.ProductVersion.Substring(0, Application.ProductVersion.LastIndexOf("."))}";

            isAutoRun = args.Contains<string>("Unattended");
            isInstall = args.Contains<string>("Install");
            if(isAutoRun | isInstall)
                this.WindowState = FormWindowState.Minimized;

            // get settings file
            string settingsPath = "";

            if(isIDE)
                // running from visual studio
                settingsPath = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Application.ExecutablePath)));
            else
                // running from executable
                settingsPath = Application.CommonAppDataPath.Remove(Application.CommonAppDataPath.LastIndexOf("."));

            settingsFile = Path.Combine(settingsPath, "Settings.xml");

            // if installing, only set permissions on settings file
            if(isInstall)
            {
                set_permissions();
                Application.Exit();
            }

            // read settings
            if(GetSettings())
            {

                try
                {
                    // get api key from file
                    string apiKeyFile = Path.Combine(settingsPath, "SendgridAPIKey.txt");
                    apiKey = File.ReadAllText(apiKeyFile);
                }
                catch(Exception ex)
                {
                    Status = $"Error getting API key: {ex.Message}";
                    LogIt.LogError(Status);
                }

                fromDate = DateTime.Today.AddDays(-fromDaysAgo);
                dtpFromDate.Value = fromDate;

                toDate = DateTime.Today.AddDays(-toDaysAgo);
                dtpToDate.Value = toDate;

                LogIt.LogInfo($"Got beginning and ending dates from settings file: FromDate={fromDate.ToShortDateString()}, ToDate={toDate.ToShortDateString()}");

                // load the schools list
                DataSet ds = GetAllSchools();
                List<string> allSchools = new List<string>();
                foreach(DataRow row in ds.Tables[0].Rows)
                    allSchools.Add(row.ItemArray[0].ToString());
                clbSchools.DataSource = allSchools;

                // check the default items
                CheckSelectedSchools();

                this.btnPreview.Enabled = true;
                this.btnProcess.Enabled = true;

                if(isAutoRun)
                    btnProcess_Click(btnProcess, null);
            }
            else
            {
                this.btnPreview.Enabled = false;
                this.btnProcess.Enabled = false;
            }
        }

        private void dtpFromDate_ValueChanged(object sender, EventArgs e)
        {
            fromDate = dtpFromDate.Value;
        }

        private void dtpToDate_ValueChanged(object sender, EventArgs e)
        {
            toDate = dtpToDate.Value;
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            process_payments("process", sender, e);
        }


        private void btnPreview_Click(object sender, EventArgs e)
        {
            process_payments("preview", sender, e);
        }

        private void lblSettings_Click(object sender, EventArgs e)
        {
            SettingsForm sf = new SettingsForm();
            sf.StartPosition = FormStartPosition.CenterParent;
            sf.ShowDialog();
            doc.Load(settingsFile);
            if(GetSettings())
            {
                fromDate = DateTime.Today.AddDays(-fromDaysAgo);
                dtpFromDate.Value = fromDate;

                toDate = DateTime.Today.AddDays(-toDaysAgo);
                dtpToDate.Value = toDate;

                LogIt.LogInfo($"Got beginning and ending dates from settings file: FromDate={fromDate.ToShortDateString()}, ToDate={toDate.ToShortDateString()}");

                CheckSelectedSchools();
                this.btnPreview.Enabled = true;
                this.btnProcess.Enabled = true;
            }
            else
            {
                this.btnPreview.Enabled = false;
                this.btnProcess.Enabled = false;
            }
        }

        #endregion

        #region methods

        private void set_permissions()
        {
            try
            {
                // Create security idenifier for all users (WorldSid)
                SecurityIdentifier sid = new SecurityIdentifier(WellKnownSidType.WorldSid, null);

                // get file info and add write, modify permissions
                FileInfo fi = new FileInfo(settingsFile);
                FileSecurity fs = fi.GetAccessControl();
                FileSystemAccessRule fsar =
                    new FileSystemAccessRule(sid, FileSystemRights.FullControl, InheritanceFlags.None, PropagationFlags.None, AccessControlType.Allow);

                fs.AddAccessRule(fsar);
                fi.SetAccessControl(fs);
                LogIt.LogInfo("Set permissions on Settings file");
            }
            catch(Exception ex)
            {
                LogIt.LogError(ex.Message);
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void process_payments(string mode, object sender, EventArgs e)
        {
            Status = null;
            TextInfo textInfo = Thread.CurrentThread.CurrentCulture.TextInfo;

            foreach(string school in clbSchools.CheckedItems)
            {
                status = $"{textInfo.ToTitleCase(mode)}ing: School={school}, From={fromDate.ToShortDateString()}, To={toDate.ToShortDateString()}";
                LogIt.LogInfo(status);
                if(Status != null)
                    Status += Environment.NewLine;
                Status += status;

                // get the payments
                DataSet ds = GetPaymentsForSchool(school);
                if(ds.Tables[0].Rows.Count == 0)
                {
                    status = $"No payments to {mode}  between {fromDate.ToShortDateString()} and {toDate.ToShortDateString()}";
                    LogIt.LogInfo(status);
                    Status += Environment.NewLine + "  " + status;
                }
                else
                {
                    int noOfStudents = ds.Tables[0].AsEnumerable().GroupBy(r => r["StudentID"]).ToList().Count;
                    status = $"Got {ds.Tables[0].Rows.Count} payments for {noOfStudents} students";
                    Status += Environment.NewLine + "  " + status;

                    if(mode == "process")
                    {
                        StringBuilder sbText = new StringBuilder();
                        StringBuilder sbHtml = new StringBuilder();
                        StringBuilder sbParentText = new StringBuilder();
                        StringBuilder sbParentHtml = new StringBuilder();

                        List<int> paymentList = new List<int>();
                        errorList = new List<string>();
                        fromEmail = from_email.Replace("{school}", school);
                        fromName = from_name.Replace("{school}", school);
                        problems = 0;

                        // loop thru table rows and build emails
                        DataTable tbl = ds.Tables[0];
                        int curStudentID = 0;
                        int curSchool_ID = 0;
                        DataColumn idCol = tbl.Columns["StudentID"];
                        foreach(DataRow row in tbl.Rows)
                        {
                            // when we encounter a new student, process the email for prior student
                            if((int)row[idCol] != curStudentID)
                            {
                                // we've hit a new student, so if not the very first student, send email
                                if(curStudentID != 0)
                                {
                                    textBody = string.Format(text_body, curStudentFirst, curStudentLast, curSchoolName, sbText.ToString(), curSchoolID);
                                    htmlBody = string.Format(html_body, curStudentFirst, curStudentLast, curSchoolName, sbHtml.ToString(), curSchoolID);

                                    try
                                    {
                                        // send the email, update payments and student if successful
                                        Response r = await SendEmail(apiKey, curStudentFirst, curStudentLast, curStudentEmail, false, fromEmail, fromName, curSchoolContact, curSchoolCcEmail, textBody, htmlBody);
                                        if(r.StatusCode == System.Net.HttpStatusCode.Accepted && paymentList.Count != 0)
                                        {
                                            status = $"  Sent email to {curStudentFirst} {curStudentLast} <{curStudentEmail}>, result={r.StatusCode.ToString()}";
                                            LogIt.LogInfo(status);
                                            Status += Environment.NewLine + "  " + status;

                                            UpdatePayments(curStudentFirst, curStudentLast, paymentList, DateTime.Now);
                                            UpdateStudent(curStudentFirst, curStudentLast, curStudentID, DateTime.Now);
                                            status = $"Updated payment records & notes for {school} student {curStudentFirst} {curStudentLast}";
                                            LogIt.LogInfo(status);
                                            Status += Environment.NewLine + "  " + status;
                                        }
                                        else if(r.StatusCode == System.Net.HttpStatusCode.PartialContent)
                                        {
                                            status = $"Missing email for {school} student {curStudentFirst} {curStudentLast}";
                                            errorList.Add(status);
                                            LogIt.LogError(status);
                                            Status += Environment.NewLine + "  " + status;
                                            problems++;
                                        }
                                        else if(r.StatusCode == System.Net.HttpStatusCode.PreconditionFailed)
                                        {
                                            status = $"Error occurred trying to send email for {school} student {curStudentFirst} {curStudentLast}: {msg}";
                                            errorList.Add(status);
                                            LogIt.LogError(status);
                                            Status += Environment.NewLine + "  " + status;
                                            problems++;
                                        }
                                    }
                                    catch(Exception ex)
                                    {
                                        status = $"Error sending email or updating tables: {ex.Message}";
                                        errorList.Add(status);
                                        LogIt.LogError(status);
                                        Status += Environment.NewLine + "  " + status;
                                        problems++;
                                    }

                                    // if any PLUS, send email to parent
                                    if(sbParentText.Length > 0)
                                    {
                                        textBody = string.Format(parent_text_body, curStudentFirst, curStudentLast, curSchoolName, sbParentText.ToString(), curSchoolID);
                                        htmlBody = string.Format(parent_html_body, curStudentFirst, curStudentLast, curSchoolName, sbParentHtml.ToString(), curSchoolID);

                                        try
                                        {
                                            // send the parent email, update student notes if successful
                                            Response r = await SendEmail(apiKey, curStudentFirst, curStudentLast, curParentEmail, true, fromEmail, fromName, curSchoolContact, curSchoolCcEmail, textBody, htmlBody);
                                            if(r.StatusCode == System.Net.HttpStatusCode.Accepted && paymentList.Count != 0)
                                            {
                                                status = $"  Sent email to parent of {curStudentFirst} {curStudentLast} <{curParentEmail}>, result={r.StatusCode.ToString()}";
                                                LogIt.LogInfo(status);
                                                Status += Environment.NewLine + "  " + status;

                                                UpdateStudent(curStudentFirst, curStudentLast, curStudentID, DateTime.Now, true);
                                                status = $"Updated notes for {school} student {curStudentFirst} {curStudentLast}";
                                                LogIt.LogInfo(status);
                                                Status += Environment.NewLine + "  " + status;
                                            }
                                            else if(r.StatusCode == System.Net.HttpStatusCode.PartialContent)
                                            {
                                                status = $"Missing email for parent(s) of {school} student {curStudentFirst} {curStudentLast}";
                                                errorList.Add(status);
                                                LogIt.LogError(status);
                                                Status += Environment.NewLine + "  " + status;
                                                problems++;
                                            }
                                            else if(r.StatusCode == System.Net.HttpStatusCode.PreconditionFailed)
                                            {
                                                status = $"Error occurred trying to send email for parent(s) of {school} student {curStudentFirst} {curStudentLast}: {msg}";
                                                errorList.Add(status);
                                                LogIt.LogError(status);
                                                Status += Environment.NewLine + "  " + status;
                                                problems++;
                                            }
                                        }
                                        catch(Exception ex)
                                        {
                                            status = $"Error sending email or updating student table: {ex.Message}";
                                            errorList.Add(status);
                                            LogIt.LogError(status);
                                            Status += Environment.NewLine + "  " + status;
                                            problems++;
                                        }
                                    }
                                }

                                // set current values, clear list of payments
                                curStudentID = (int)row[tbl.Columns["StudentID"]];
                                curSchool_ID = (int)row[tbl.Columns["School_ID"]];
                                curSchoolID = (row[tbl.Columns["SchoolID"]] ?? "").ToString();
                                curSchoolName = (row[tbl.Columns["SchoolName"]] ?? "").ToString();
                                curSchoolContact = "Financial Aid Department"; // was this before: (row[tbl.Columns["ContactName"]] ?? "").ToString();
                                curSchoolContactEmail = (row[tbl.Columns["ContactEmail"]] ?? "").ToString();
                                curSchoolCcEmail = (row[tbl.Columns["EmailForPaymentAdvice"]] ?? "").ToString();
                                curSchoolErrorEmail = (row[tbl.Columns["EmailForPaymentAdviceErrors"]] ?? "").ToString();
                                curStudentFirst = (row[tbl.Columns["FirstName"]] ?? "").ToString();
                                curStudentLast = (row[tbl.Columns["LastName"]] ?? "").ToString();
                                curStudentEmail = (row[tbl.Columns["Email"]] ?? "").ToString();
                                curParentEmail = (row[tbl.Columns["ParentEmail"]] ?? "").ToString();
                                sbText = new StringBuilder();
                                sbHtml = new StringBuilder();
                                sbParentText = new StringBuilder();
                                sbParentHtml = new StringBuilder();
                                paymentList = new List<int>();
                            }

                            // append to list of payment IDs
                            paymentList.Add((int)row[tbl.Columns["PaymentID"]]);

                            // save payment details for text and html message
                            SavePaymentDetails(tbl, row, ref sbText);
                            SavePaymentDetails(tbl, row, ref sbHtml);

                            // if this is PLUS loan, keep track of it so we can send another email to parent
                            if((string)row[tbl.Columns["FedPgmName"]] == "DL PLUS")
                            {
                                SavePaymentDetails(tbl, row, ref sbParentText);
                                SavePaymentDetails(tbl, row, ref sbParentHtml);
                            }
                        }

                        // send email and update if any payment details accumulated for last student in school.
                        if(sbText.Length != 0)
                        {
                            textBody = string.Format(text_body, curStudentFirst, curStudentLast, curSchoolName, sbText.ToString(), curSchoolID);
                            htmlBody = string.Format(html_body, curStudentFirst, curStudentLast, curSchoolName, sbHtml.ToString(), curSchoolID);

                            try
                            {
                                Response r = await SendEmail(apiKey, curStudentFirst, curStudentLast, curStudentEmail, false, fromEmail, fromName, curSchoolContact, curSchoolCcEmail, textBody, htmlBody);
                                if(r.StatusCode == System.Net.HttpStatusCode.Accepted && paymentList.Count != 0)
                                {
                                    status = $"  Sent email to {curStudentFirst} {curStudentLast} <{curStudentEmail}>, result={r.StatusCode.ToString()}";
                                    LogIt.LogInfo(status);
                                    Status += Environment.NewLine + "  " + status;

                                    UpdatePayments(curStudentFirst, curStudentLast, paymentList, DateTime.Now);
                                    UpdateStudent(curStudentFirst, curStudentLast, curStudentID, DateTime.Now);
                                    status = $"Updated payment records & notes for student {curStudentFirst} {curStudentLast}";
                                    LogIt.LogInfo(status);
                                    Status += Environment.NewLine + "  " + status;
                                }
                                else if(r.StatusCode == System.Net.HttpStatusCode.PartialContent)
                                {
                                    status = $"Missing email for {school} student {curStudentFirst} {curStudentLast}";
                                    errorList.Add(status);
                                    LogIt.LogError(status);
                                    Status += Environment.NewLine + "  " + status;
                                    problems++;
                                }
                                else if(r.StatusCode == System.Net.HttpStatusCode.PreconditionFailed)
                                {
                                    status = $"Error occurred trying to send email for {school} student {curStudentFirst} {curStudentLast}: {msg}";
                                    errorList.Add(status);
                                    LogIt.LogError(status);
                                    Status += Environment.NewLine + "  " + status;
                                    problems++;
                                }
                            }
                            catch(Exception ex)
                            {
                                status = $"Error sending email or updating tables: {ex.Message}";
                                errorList.Add(status);
                                LogIt.LogError(status);
                                Status += Environment.NewLine + "  " + status;
                                problems++;
                            }

                            // if any PLUS payments, send email to parent
                            if(sbParentText.Length > 0)
                            {
                                textBody = string.Format(parent_text_body, curStudentFirst, curStudentLast, curSchoolName, sbParentText.ToString(), curSchoolID);
                                htmlBody = string.Format(parent_html_body, curStudentFirst, curStudentLast, curSchoolName, sbParentHtml.ToString(), curSchoolID);

                                try
                                {
                                    // send the parent email, update student notes if successful
                                    Response r = await SendEmail(apiKey, curStudentFirst, curStudentLast, curParentEmail, true, fromEmail, fromName, curSchoolContact, curSchoolCcEmail, textBody, htmlBody);
                                    if(r.StatusCode == System.Net.HttpStatusCode.Accepted && paymentList.Count != 0)
                                    {
                                        status = $"  Sent email to parent of {curStudentFirst} {curStudentLast} <{curParentEmail}>, result={r.StatusCode.ToString()}";
                                        LogIt.LogInfo(status);
                                        Status += Environment.NewLine + "  " + status;

                                        UpdateStudent(curStudentFirst, curStudentLast, curStudentID, DateTime.Now, true);
                                        status = $"Updated notes for {school} student {curStudentFirst} {curStudentLast}";
                                        LogIt.LogInfo(status);
                                        Status += Environment.NewLine + "  " + status;
                                    }
                                    else if(r.StatusCode == System.Net.HttpStatusCode.PartialContent)
                                    {
                                        status = $"Missing email for parent(s) of {school} student {curStudentFirst} {curStudentLast}";
                                        errorList.Add(status);
                                        LogIt.LogError(status);
                                        Status += Environment.NewLine + "  " + status;
                                        problems++;
                                    }
                                    else if(r.StatusCode == System.Net.HttpStatusCode.PreconditionFailed)
                                    {
                                        status = $"Error occurred trying to send email for parent(s) of {school} student {curStudentFirst} {curStudentLast}: {msg}";
                                        errorList.Add(status);
                                        LogIt.LogError(status);
                                        Status += Environment.NewLine + "  " + status;
                                        problems++;
                                    }
                                }
                                catch(Exception ex)
                                {
                                    status = $"Error sending email or updating student table: {ex.Message}";
                                    errorList.Add(status);
                                    LogIt.LogError(status);
                                    Status += Environment.NewLine + "  " + status;
                                    problems++;
                                }
                            }
                        }

                        string errMsg = (problems > 0) ? "errors" : "no errors";
                        status = $"Processing complete for {school} with {errMsg}.";
                        LogIt.LogInfo(status);
                        Status += Environment.NewLine + "  " + status;

                        // done with school, send email(s) if any errors occurred
                        if(mode == "process" && ERRsTo.Count > 0 && errorList.Count > 0)
                        {
                            textBody = string.Format(text_error_body, curSchoolID, string.Join("\n", errorList));
                            htmlBody = string.Format(html_error_body, curSchoolID, string.Join("\n", errorList));
                            try
                            {
                                Response r = await SendErrorEmail(apiKey, ERRsTo, errorList, curSchoolContact, curSchoolErrorEmail, textBody, htmlBody);
                                if(r.StatusCode == System.Net.HttpStatusCode.Accepted)
                                {
                                    status = $"Sent error email(s) for {school} to {ERRsTo}, result={r.StatusCode.ToString()}";
                                    LogIt.LogInfo(status);
                                    Status += Environment.NewLine + "  " + status;
                                }
                                else if(r.StatusCode == System.Net.HttpStatusCode.PartialContent)
                                {
                                    status = $"Missing email address for sending error emails";
                                    LogIt.LogError(status);
                                    Status += Environment.NewLine + "  " + status;
                                }
                                else if(r.StatusCode == System.Net.HttpStatusCode.PreconditionFailed)
                                {
                                    status = $"Error occurred trying to send error email(s) for {school}: {msg}";
                                    LogIt.LogError(status);
                                    Status += Environment.NewLine + "  " + status;
                                }
                            }
                            catch(Exception ex)
                            {
                                status = $"Error sending error email(s): {ex.Message}";
                                LogIt.LogError(status);
                                Status += Environment.NewLine + "  " + status;
                            }
                        }
                    }
                }
            }
            if(isAutoRun)
                Application.Exit();
        }

        private void SavePaymentDetails(DataTable tbl, DataRow row, ref StringBuilder sb)
        {
            if(sb.Length != 0)
                sb.Append("\n");
            sb.Append("Check #: ");
            sb.Append(row[tbl.Columns["CkNo"]].ToString());
            sb.Append("   Date: ");
            sb.Append(string.Format("{0:M/d/yyyy}", row[tbl.Columns["CkDate"]]));
            sb.Append("   Federal Program: ");
            sb.Append((string)row[tbl.Columns["FedPgmName"]]);
            sb.Append("   Amount: ");
            sb.Append(row[tbl.Columns["Amount"]].ToString());
        }

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
                txtStatus.SelectionStart = txtStatus.TextLength;
                txtStatus.ScrollToCaret();
            }
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

        private bool GetSettings()
        {
            try
            {
                doc = new ConfigXmlDocument();
                doc.Load(settingsFile);
                sendTo = GetSetting(doc, "SendTo");
                sendParentTo = GetSetting(doc, "SendParentTo");

                var tempCCs = GetSetting(doc, "SendCC");
                if(tempCCs != "")
                    CCsTo = tempCCs.Split(',').ToList<string>();

                var tempBCCs = GetSetting(doc, "SendBCC");
                if(tempBCCs != "")
                    BCCsTo = tempBCCs.Split(',').ToList<string>();

                var tempERRs = GetSetting(doc, "SendErrors");
                if(tempERRs != "")
                    ERRsTo = tempERRs.Split(',').ToList<string>();

                var tempFromDaysAgo = GetSetting(doc, "FromDaysAgo");
                if(!int.TryParse(tempFromDaysAgo, out fromDaysAgo))
                    fromDaysAgo = 12;

                var tempToDaysAgo = GetSetting(doc, "ToDaysAgo");
                if(!int.TryParse(tempToDaysAgo, out toDaysAgo))
                    toDaysAgo = 5;

                from_email = GetSetting(doc, "FromEmail");
                from_name = GetSetting(doc, "FromName");
                return true;
            }
            catch(Exception ex)
            {
                Status = $"Error getting application settings: {ex.Message}";
                LogIt.LogError(Status);
                return false;
            }

        }

        private string GetSetting(XmlDocument doc, string settingName)
        {
            string response = "";
            try
            {
                response = ((XmlElement)doc.SelectSingleNode($"/Settings/setting[@name='{settingName}']")).GetAttribute("value");

            }
            catch(Exception)
            {
            }
            return response;
        }

        private DataSet GetAllSchools()
        {
            DataSet ds = new DataSet();
            try
            {
                LogIt.LogInfo("Getting schools.");
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

            }
            catch(Exception ex)
            {
                var msg = $"Error getting schools: {ex.Message}";
                LogIt.LogError(msg);
            }
            return ds;
        }

        private void UpdateStudent(string studentFirst, string studentLast, int student, DateTime now, bool isParent = false)
        {
            string note = "EFT email sent to " + (isParent ? "parent" : "student");
            string update_sql = $"Update Student set Notes = concat(Notes, '{"\r\n"}', '{now.ToString("MM/dd/yyyy")}: {note}.') where StudentID = ({student});";

            using(ODBCClass o = new ODBCClass("MySql"))
            {
                using(OdbcCommand oCommand = o.GetCommand(update_sql))
                {
                    oCommand.CommandText = update_sql;
                    try
                    {
                        var rows = oCommand.ExecuteNonQuery();
                        LogIt.LogInfo($"Updated notes for {studentFirst} {studentLast} (student ID {student}), Rows affected: {rows}");
                    }
                    catch(Exception ex)
                    {
                        LogIt.LogError($"Could not update student {student}: {ex.Message}");
                    }
                }
            }
        }

        private void UpdatePayments(string studentFirst, string studentLast, List<int> paymentList, DateTime now)
        {
            string inList = string.Join(",", paymentList.ConvertAll(Convert.ToString));
            string update_sql = $"Update Payments p set p.AdviceDate = '{now.ToString("yyyy-MM-dd HH:mm:ss")}' where p.PaymentID in({inList});";

            using(ODBCClass o = new ODBCClass("MySql"))
            {
                using(OdbcCommand oCommand = o.GetCommand(update_sql))
                {
                    oCommand.CommandText = update_sql;
                    oCommand.Parameters.AddWithValue("@date", now.ToString("yyyy-MM-dd HH:mm:ss"));
                    oCommand.Parameters.AddWithValue("@pymts", inList);
                    try
                    {
                        var rows = oCommand.ExecuteNonQuery();
                        LogIt.LogInfo($"Updated payments for {studentFirst} {studentLast}, IDs {inList}, Rows affected: {rows}");
                    }
                    catch(Exception ex)
                    {
                        LogIt.LogError($"Could not update payments: {ex.Message}");
                    }
                }
            }
        }

        private DataSet GetPaymentsForSchool(string school)
        {
            DataSet ds = new DataSet();
            try
            {
                //if(Status != null)
                //    Status += Environment.NewLine;
                status = $"Getting payments made";
                LogIt.LogInfo(status);
                Status += Environment.NewLine + "  " + status;

                using(ODBCClass o = new ODBCClass("MySql"))
                {
                    var procName = "payments_made(?,?,?)";
                    var parms = new KeyValuePair<string, string>[]{
                                    new KeyValuePair<string,string>("_school", school),
                                    new KeyValuePair<string,string>("_fromDate", fromDate.ToString("yyyy-MM-dd")),
                                    new KeyValuePair<string,string>("_toDate", toDate.ToString("yyyy-MM-dd"))
                                };

                    using(OdbcCommand oCommand = o.GetCommand(procName, parms))
                    {
                        using(OdbcDataAdapter oAdapter = new OdbcDataAdapter(oCommand))
                        {
                            oAdapter.Fill(ds);
                        }
                    }

                }
            }
            catch(Exception)
            {
            }
            return ds;
        }

        private async Task<Response> SendEmail(string apiKey, string studFirst, string studLast, string emailAddress, bool isParent, string fromEmail, string fromName,
                                               string schoolContact, string schoolEmail, string txtBody, string htmBody)
        {
            return await Task<Response>.Run(async () =>
             {
                 Response resp = new Response(System.Net.HttpStatusCode.PartialContent, null, null);
                 var client = new SendGridClient(apiKey);
                 var from = new EmailAddress(fromEmail, fromName);
                 var subject = "Financial Aid Awarded";
                 var toName = (isParent ? "Parent(s) of " : "") + $"{studFirst} {studLast}";

                 EmailAddress to;
                 if(isParent)
                     to = new EmailAddress(sendParentTo == "{parent}" ? emailAddress : sendParentTo, toName);
                 else
                     to = new EmailAddress(sendTo == "{student}" ? emailAddress : sendTo, toName);
                 var plainTextContent = txtBody;
                 var htmlContent = htmBody;

                 if(to.Email == "")
                 {
                 }
                 else
                 {
                     List<string> schoolEmails = schoolEmail.Split(',').ToList<string>();
                     try
                     {
                         var sgMsg = MailHelper.CreateSingleEmail(from, to, subject, plainTextContent, htmlContent);

                         // add any CCs from settings
                         if(CCsTo.Count > 0)
                         {
                             var ccList = new List<EmailAddress>();
                             foreach(var cc in CCsTo)
                             {
                                 if(cc == "{school}")
                                 {
                                     foreach(var email in schoolEmails)
                                     {
                                         ccList.Add(new EmailAddress(email, "Financial Aid Department"));
                                     }
                                 }
                                 else
                                     ccList.Add(new EmailAddress(cc, "EmailPaymentAdvice Administrator"));
                             }
                             sgMsg.AddCcs(ccList);
                         }

                         // add any BCCs from settings
                         if(BCCsTo.Count > 0)
                         {
                             var bccList = new List<EmailAddress>();
                             foreach(var bcc in BCCsTo)
                             {
                                 if(bcc == "{school}")
                                 {
                                     foreach(var email in schoolEmails)
                                     {
                                         bccList.Add(new EmailAddress(email, "Financial Aid Department"));
                                     }
                                 }
                                 else
                                     bccList.Add(new EmailAddress(bcc, "EmailPaymentAdvice Administrator"));
                             }
                             sgMsg.AddBccs(bccList);
                         }
                         resp = await client.SendEmailAsync(sgMsg);
                     }
                     catch(Exception ex)
                     {
                         msg = ex.Message;
                         resp.StatusCode = System.Net.HttpStatusCode.PreconditionFailed;
                     }
                 }
                 return resp;
             });
        }

        private async Task<Response> SendErrorEmail(string apiKey, List<string> ERRsTo, List<string> errorList, string schoolContact, string schoolErrorEmail,
                                                    string txtBody, string htmBody)
        {
            return await Task<Response>.Run(async () =>
            {
                Response resp = new Response(System.Net.HttpStatusCode.PartialContent, null, null);
                var client = new SendGridClient(apiKey);
                var from = new EmailAddress("Info@ShamrocksFA.com", "EmailPaymentAdvice");
                var subject = "Errors Sending Payment Advice Emails";

                // create new list of emails & names, and resolve {school} emails
                List<string> emails = new List<string>();
                List<string> emailNames = new List<string>();
                List<string> schoolEmails = schoolErrorEmail.Split(',').ToList<string>();
                foreach(var errTo in ERRsTo)
                {
                    if(errTo == "{school}")
                    {
                        foreach(var schoolEmail in schoolEmails)
                        {
                            emails.Add(schoolEmail);
                            emailNames.Add("Financial Aid Department");
                        }
                    }
                    else
                    {
                        emails.Add(errTo);
                        emailNames.Add("EmailPaymentAdvice Administrator");
                    }
                }

                var plainTextContent = txtBody;
                var htmlContent = htmBody;

                // get TO email address from first item in lists, then remove from lists
                var to = new EmailAddress(emails[0], emailNames[0]);
                emails.RemoveAt(0);
                emailNames.RemoveAt(0);

                try
                {
                    // create email & add any remaining email addresses as CCs
                    var msg = MailHelper.CreateSingleEmail(from, to, subject, plainTextContent, htmlContent);
                    if(emails.Count > 0)
                    {
                        var ccList = new List<EmailAddress>();
                        for(int i = 0; i < emails.Count; i++)
                        {
                            ccList.Add(new EmailAddress(emails[i], emailNames[i]));
                        }
                        msg.AddCcs(ccList);
                    }

                    resp = await client.SendEmailAsync(msg);
                }
                catch(Exception ex)
                {
                    msg = ex.Message;
                    resp.StatusCode = System.Net.HttpStatusCode.PreconditionFailed;
                }

                return resp;
            });
        }

        #endregion

    }

}

