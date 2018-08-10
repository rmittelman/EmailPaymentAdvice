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

namespace EmailPaymentAdvice
{
    public partial class Form1 : Form
    {
        private List<string> CCsTo = new List<string>();
        private List<string> BCCsTo = new List<string>();
        private List<string> ERRsTo = new List<string>();
        private List<string> errorList = new List<string>();
        private string sendTo;
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
        private string settingsFile;
        private XmlDocument doc;

        private string text_body = "Dear {0} {1},\n\nYour Financial Aid to attend {2} has been posted "
                    + "to your student account as follows:\n\n{3}\n\n"
                    + "Please understand that as a student you have the right to cancel all or a portion of a loan "
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

        private string text_error_body = "Dear EmailPaymentAdvice Administrator,\n\nThe following errors have occurred "
                    + "while processing emails for {0}:\n\n{1}\n\n"
                    + "For further details, please review the log file at \"M:\\EmailPaymentAdvice\\Logs\\App.log\".\n\n"
                    + "Please do not reply to this email because this email box is not monitored.";

        private string html_error_body = "<p>Dear EmailPaymentAdvice Administrator,</p><p>The following errors have occurred "
                    + "while processing emails for {0}:</p><p><pre>{1}</pre></p>"
                    + "<p>For further details, please review the log file at <strong>M:\\EmailPaymentAdvice\\Logs\\App.log</strong>.</p>"
                    + "<p>Please do not reply to this email because this email box is not monitored.</p>";

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

            isAutoRun = args.Contains<string>("Unattended");
            if(isAutoRun)
                this.WindowState = FormWindowState.Minimized;

            // setup ToolTips
            ToolTip toolTip1 = new ToolTip();
            toolTip1.AutoPopDelay = 5000;
            toolTip1.InitialDelay = 1000;
            toolTip1.ReshowDelay = 500;
            toolTip1.ShowAlways = true;
            toolTip1.SetToolTip(this.lblSettings, "Settings");

            lblVersion.Text = $"v {Application.ProductVersion.Substring(0, Application.ProductVersion.LastIndexOf("."))}";

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
                sendTo = GetSetting(doc, "SendTo");

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

                // get api key from file
                string apiKeyFile = Path.Combine(settingsPath, "SendgridAPIKey.txt");
                apiKey = File.ReadAllText(apiKeyFile);
            }
            catch(Exception ex)
            {
                Status = $"Error getting application settings: {ex.Message}";
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

            if(isAutoRun)
                btnProcess_Click(btnProcess, null);
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
            CheckSelectedSchools();
        }

        #endregion

        #region methods

        private async void process_payments(string mode, object sender, EventArgs e)
        {
            Status = null;
            TextInfo textInfo = Thread.CurrentThread.CurrentCulture.TextInfo;

            foreach(string school in clbSchools.CheckedItems)
            {
                //StringBuilder sbText = new StringBuilder();
                //StringBuilder sbHtml = new StringBuilder();
                //List<int> paymentList = new List<int>();
                //errorList = new List<string>();
                //fromEmail = from_email.Replace("{school}", school);
                //fromName = from_name.Replace("{school}", school);
                //Status = $"Processing: School={school}, From={fromDate.ToShortDateString()}, To={toDate.ToShortDateString()}";
                //LogIt.LogInfo(Status);
                //int problems = 0;

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
                        List<int> paymentList = new List<int>();
                        errorList = new List<string>();
                        fromEmail = from_email.Replace("{school}", school);
                        fromName = from_name.Replace("{school}", school);
                        int problems = 0;

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
                                        Response r = await SendEmail(apiKey, curStudentFirst, curStudentLast, curStudentEmail, fromEmail, fromName, curSchoolContact, curSchoolContactEmail, textBody, htmlBody);
                                        if(r.StatusCode == System.Net.HttpStatusCode.Accepted && paymentList.Count != 0)
                                        {
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
                                            status = $"Error occurred trying to send email for {school} student {curStudentFirst} {curStudentLast}";
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

                                }

                                // set current values, clear list of payments
                                curStudentID = (int)row[tbl.Columns["StudentID"]];
                                curSchool_ID = (int)row[tbl.Columns["School_ID"]];
                                curSchoolID = (row[tbl.Columns["SchoolID"]] ?? "").ToString();
                                curSchoolName = (row[tbl.Columns["SchoolName"]] ?? "").ToString();
                                curSchoolContact = (row[tbl.Columns["ContactName"]] ?? "").ToString();
                                curSchoolContactEmail = (row[tbl.Columns["ContactEmail"]] ?? "").ToString();
                                curStudentFirst = (row[tbl.Columns["FirstName"]] ?? "").ToString();
                                curStudentLast = (row[tbl.Columns["LastName"]] ?? "").ToString();
                                curStudentEmail = (row[tbl.Columns["Email"]] ?? "").ToString();
                                sbText = new StringBuilder();
                                sbHtml = new StringBuilder();
                                paymentList = new List<int>();

                            }

                            // append to list of payments and stringBuilders
                            paymentList.Add((int)row[tbl.Columns["PaymentID"]]);

                            if(sbText.Length != 0)
                                sbText.Append("\n");
                            sbText.Append("Check #: ");
                            sbText.Append(row[tbl.Columns["CkNo"]].ToString());
                            sbText.Append("   Date: ");
                            sbText.Append(string.Format("{0:M/d/yyyy}", row[tbl.Columns["CkDate"]]));
                            sbText.Append("   Federal Program: ");
                            sbText.Append((string)row[tbl.Columns["FedPgmName"]]);
                            sbText.Append("   Amount: ");
                            sbText.Append(row[tbl.Columns["Amount"]].ToString());

                            if(sbHtml.Length != 0)
                                sbHtml.Append("\n");
                            sbHtml.Append("Check #: ");
                            sbHtml.Append(row[tbl.Columns["CkNo"]].ToString());
                            sbHtml.Append("   Date: ");
                            sbHtml.Append(string.Format("{0:M/d/yyyy}", row[tbl.Columns["CkDate"]]));
                            sbHtml.Append("   Federal Program: ");
                            sbHtml.Append((string)row[tbl.Columns["FedPgmName"]]);
                            sbHtml.Append("   Amount: ");
                            sbHtml.Append(row[tbl.Columns["Amount"]].ToString());

                        }

                        // send email and update if any payment details accumulated for last student in school.
                        if(sbText.Length != 0)
                        {
                            textBody = string.Format(text_body, curStudentFirst, curStudentLast, curSchoolName, sbText.ToString(), curSchoolID);
                            htmlBody = string.Format(html_body, curStudentFirst, curStudentLast, curSchoolName, sbHtml.ToString(), curSchoolID);

                            try
                            {
                                Response r = await SendEmail(apiKey, curStudentFirst, curStudentLast, curStudentEmail, fromEmail, fromName, curSchoolContact, curSchoolContactEmail, textBody, htmlBody);
                                if(r.StatusCode == System.Net.HttpStatusCode.Accepted && paymentList.Count != 0)
                                {
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
                                    status = $"Error occurred trying to send email for {school} student {curStudentFirst} {curStudentLast}";
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
                                Response r = await SendErrorEmail(apiKey, ERRsTo, errorList, textBody, htmlBody);
                                if(r.StatusCode == System.Net.HttpStatusCode.Accepted)
                                {
                                    status = $"Sent error email(s) for {school}";
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
                                    status = $"Error occurred trying to send error email(s) for {school}";
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

        private void CheckSelectedSchools()
        {
            // first un-check all
            for(int i = 0; i < clbSchools.Items.Count; i++)
                clbSchools.SetItemChecked(i, false);

            // get list of desired schools from settings and check those
            string tempSchools = GetSetting(doc, "Schools");
            List<string> selectedSchools = tempSchools.Split(',').ToList<string>();
            for(int i = 0; i < selectedSchools.Count; i++)
                clbSchools.SetItemChecked(clbSchools.Items.IndexOf(selectedSchools[i]), true);
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

        private void UpdateStudent(string studentFirst, string studentLast, int student, DateTime now)
        {
            string update_sql = $"Update Student set Notes = concat(Notes, '{"\r\n"}', '{now.ToString("MM/dd/yyyy")}: EFT email sent.') where StudentID = ({student});";

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

        private async Task<Response> SendEmail(string apiKey, string studFirst, string studLast, string emailAddress, string fromEmail, string fromName, string schoolContact, string schoolEmail, string txtBody, string htmBody)
        {
            return await Task<Response>.Run(async () =>
             {
                 Response resp = new Response(System.Net.HttpStatusCode.PartialContent, null, null);
                 var client = new SendGridClient(apiKey);
                 var from = new EmailAddress(fromEmail, fromName);
                 var subject = "Financial Aid Awarded";
                 var to = new EmailAddress(sendTo == "{student}" ? emailAddress : sendTo, $"{studFirst} {studLast}");
                 var plainTextContent = txtBody;
                 var htmlContent = htmBody;

                 if(to.Email == "")
                 {
                     status = $"Missing email for {fromName} student {studFirst} {studLast}";
                     errorList.Add(status);
                     LogIt.LogError(status);
                     Status += Environment.NewLine + "  " + status;
                 }
                 else
                 {
                     try
                     {
                         var msg = MailHelper.CreateSingleEmail(from, to, subject, plainTextContent, htmlContent);

                         // add any CCs from settings
                         if(CCsTo.Count > 0)
                         {
                             var ccList = new List<EmailAddress>();
                             foreach(var cc in CCsTo)
                             {
                                 ccList.Add(new EmailAddress(cc == "{school}" ? schoolEmail : cc, $"{studFirst} {studLast}"));
                             }
                             msg.AddCcs(ccList);
                         }

                         // add any BCCs from settings
                         if(BCCsTo.Count > 0)
                         {
                             var bccList = new List<EmailAddress>();
                             foreach(var bcc in BCCsTo)
                             {
                                 bccList.Add(new EmailAddress(bcc == "{school}" ? schoolEmail : bcc, $"{studFirst} {studLast}"));
                             }
                             msg.AddBccs(bccList);
                         }
                         resp = await client.SendEmailAsync(msg);

                         status = $"  Sent email to {to.Name} <{to.Email}>, result={resp.StatusCode.ToString()}";
                         LogIt.LogInfo(status);
                         Status += Environment.NewLine + "  " + status;
                     }
                     catch(Exception ex)
                     {
                         status = $"Could not send email: {ex.Message}";
                         LogIt.LogError(status);
                         Status += Environment.NewLine + "  " + status;
                         resp.StatusCode = System.Net.HttpStatusCode.PreconditionFailed;
                     }
                 }
                 return resp;
             });
        }

        private async Task<Response> SendErrorEmail(string apiKey, List<string> ERRsTo, List<string> errorList, string txtBody, string htmBody)
        {
            return await Task<Response>.Run(async () =>
            {
                Response resp = new Response(System.Net.HttpStatusCode.PartialContent, null, null);
                var client = new SendGridClient(apiKey);
                var from = new EmailAddress("Info@ShamrocksFA.com", "EmailPaymentAdvice");
                var subject = "Errors Occurred";
                var to = new EmailAddress(ERRsTo[0], "EmailPaymentAdvice Administrator");
                var plainTextContent = txtBody;
                var htmlContent = htmBody;

                // delete first address, we already used that
                ERRsTo.RemoveAt(0);

                try
                {
                    var msg = MailHelper.CreateSingleEmail(from, to, subject, plainTextContent, htmlContent);

                    // add any emails remaining as CCs
                    if(ERRsTo.Count > 0)
                    {
                        var ccList = new List<EmailAddress>();
                        foreach(var cc in ERRsTo)
                        {
                            ccList.Add(new EmailAddress(cc, "EmailPaymentAdvice Application"));
                        }
                        msg.AddCcs(ccList);
                    }

                    resp = await client.SendEmailAsync(msg);

                    status = $"Sent email to {to.Name} <{to.Email}>, result={resp.StatusCode.ToString()}";
                    LogIt.LogInfo(status);
                    Status += Environment.NewLine + "  " + status;
                }
                catch(Exception ex)
                {
                    status = $"Could not send email: {ex.Message}";
                    LogIt.LogError(status);
                    Status += Environment.NewLine + "  " + status;
                    resp.StatusCode = System.Net.HttpStatusCode.PreconditionFailed;
                }

                return resp;
            });
        }

        #endregion

    }

}

