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


namespace EmailPaymentAdvice
{
    public partial class Form1 : Form
    {
        private List<string> CCs = new List<string>();
        private List<string> Schools = new List<string>();
        private string sendTo;
        private string bccTo;
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
        string from_email;
        string from_name;
        string fromEmail;
        string fromName;
        private string msg = "";
        private string tempSchools = "";
        private string tempCCs = "";
        private bool isIDE = (Debugger.IsAttached == true);
        string[] args = Environment.GetCommandLineArgs();
        private bool isAutoRun;

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
            if (isAutoRun)
                this.WindowState = FormWindowState.Minimized;

            lblVersion.Text = $"v {Application.ProductVersion.Substring(0, Application.ProductVersion.LastIndexOf("."))}";
            
            // get settings
            try
            {
                string settingsPath = "";
                if (isIDE)
                    settingsPath = Path.GetDirectoryName(Application.ExecutablePath);
                else
                    settingsPath = Application.CommonAppDataPath.Remove(Application.CommonAppDataPath.LastIndexOf("."));

                string settingsFile = Path.Combine(settingsPath, "Settings.xml");
                XmlDocument doc = new ConfigXmlDocument();
                doc.Load(settingsFile);
                apiKey = GetSetting(doc, "SENDGRID_API_KEY");
                sendTo = GetSetting(doc, "SendTo");
                tempCCs = GetSetting(doc, "SendCC");
                bccTo = GetSetting(doc, "SendBCC");
                tempSchools = GetSetting(doc, "Schools");
                var tempFromDaysAgo = GetSetting(doc, "FromDaysAgo");
                var tempToDaysAgo = GetSetting(doc, "ToDaysAgo");
                from_email = GetSetting(doc, "FromEmail");
                from_name = GetSetting(doc, "FromName");

                doc = null;

                if (!int.TryParse(tempFromDaysAgo, out fromDaysAgo))
                    fromDaysAgo = 7;

                if (!int.TryParse(tempToDaysAgo, out toDaysAgo))
                    toDaysAgo = 14;
            }
            catch (Exception ex)
            {
                Status = $"Error getting application settings: {ex.Message}";
                LogIt.LogError(Status);
            }

            fromDate = DateTime.Today.AddDays(-fromDaysAgo);
            dtpFromDate.Value = fromDate;
            toDate = DateTime.Today.AddDays(-toDaysAgo);
            dtpToDate.Value = toDate;
            LogIt.LogInfo($"Got beginning and ending dates from settings file: FromDate={fromDate.ToShortDateString()}, ToDate={toDate.ToShortDateString()}");

            Schools = tempSchools.Split(',').ToList<string>();
            if(tempCCs != "")
                CCs = tempCCs.Split(',').ToList<string>();

            // load the schools list
            DataSet ds = GetAllSchools();
            List<string> allSchools = new List<string>();
            foreach (DataRow row in ds.Tables[0].Rows)
            {
                allSchools.Add(row.ItemArray[0].ToString());
            }
            clbSchools.DataSource = allSchools;

            // check the items
            for (int i = 0; i < Schools.Count; i++)
            {
                clbSchools.SetItemChecked(clbSchools.Items.IndexOf(Schools[i]), true);
            }
            if (isAutoRun)
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

        private async void btnProcess_Click(object sender, EventArgs e)
        {


            foreach (string school in clbSchools.CheckedItems)
            {
                StringBuilder sbText = new StringBuilder();
                StringBuilder sbHtml = new StringBuilder();
                List<int> paymentList = new List<int>();
                fromEmail = from_email.Replace("{school}", school);
                fromName = from_name.Replace("{school}", school);
                Status = $"Processing: School={school}, From={fromDate.ToShortDateString()}, To={toDate.ToShortDateString()}";
                LogIt.LogInfo(Status);

                // get the payments
                DataSet ds = GetPaymentsForSchool(school);
                if (ds.Tables[0].Rows.Count == 0)
                {
                    Status = $"No payments to process for {school} between {fromDate.ToShortDateString()} and {toDate.ToShortDateString()}";
                    LogIt.LogInfo(Status);
                }
                else
                {
                    // loop thru table rows and build emails
                    DataTable tbl = ds.Tables[0];
                    int curStudentID = 0;
                    int curSchool_ID = 0;
                    DataColumn idCol = tbl.Columns["StudentID"];
                    foreach (DataRow row in tbl.Rows)
                    {
                        // when we encounter a new student, process the email for prior student
                        if ((int)row[idCol] != curStudentID)
                        {
                            // we've hit a new student, so if not the very first student, send email
                            if (curStudentID != 0)
                            {
                                textBody = string.Format(text_body, curStudentFirst, curStudentLast, curSchoolName, sbText.ToString(), curSchoolID);
                                htmlBody = string.Format(html_body, curStudentFirst, curStudentLast, curSchoolName, sbHtml.ToString(), curSchoolID);

                                try
                                {
                                    // send the email, update payments and student if successful
                                    Response r = await SendEmail(apiKey, curStudentFirst, curStudentLast, curStudentEmail, fromEmail, fromName, curSchoolContact, curSchoolContactEmail, textBody, htmlBody);
                                    if (r.StatusCode == System.Net.HttpStatusCode.Accepted && paymentList.Count != 0)
                                    {
                                        UpdatePayments(curStudentFirst, curStudentLast, paymentList, DateTime.Now);
                                        UpdateStudent(curStudentFirst, curStudentLast, curStudentID, DateTime.Now);
                                        Status = $"Updated payment records & notes for {curStudentFirst} {curStudentLast}";
                                    }
                                    //else if (r.StatusCode == System.Net.HttpStatusCode.PartialContent)
                                    //{
                                    //    Status = $"Missing email for {curStudentFirst} {curStudentLast}";
                                    //}
                                    //else if (r.StatusCode == System.Net.HttpStatusCode.PreconditionFailed)
                                    //{
                                    //    Status = $"Error occurred trying to send email for {curStudentFirst} {curStudentLast}";
                                    //}
                                }
                                catch (Exception ex)
                                {
                                    Status = $"Error sending email or updating tables: {ex.Message}";
                                    LogIt.LogError(Status);
                                }

                            }

                            // set current values, clear list of payments
                            curStudentID = (int)row[tbl.Columns["StudentID"]];
                            curSchool_ID = (int)row[tbl.Columns["School_ID"]];
                            curSchoolID = (string)row[tbl.Columns["SchoolID"]] ?? "";
                            curSchoolName = (string)row[tbl.Columns["SchoolName"]] ?? "";
                            curSchoolContact = (string)row[tbl.Columns["ContactName"]] ?? "";
                            curSchoolContactEmail = (string)row[tbl.Columns["ContactEmail"]] ?? "";
                            curStudentFirst = (string)row[tbl.Columns["FirstName"]] ?? "";
                            curStudentLast = (string)row[tbl.Columns["LastName"]] ?? "";
                            curStudentEmail = (string)row[tbl.Columns["Email"]] ?? "";

                            sbText = new StringBuilder();
                            sbHtml = new StringBuilder();
                            paymentList = new List<int>();

                        }

                        // append to list of payments and stringBuilders
                        paymentList.Add((int)row[tbl.Columns["PaymentID"]]);

                        if (sbText.Length != 0)
                            sbText.Append("\n");
                        sbText.Append("Check #: ");
                        sbText.Append(row[tbl.Columns["CkNo"]].ToString());
                        sbText.Append("   Date: ");
                        sbText.Append(string.Format("{0:M/d/yyyy}", row[tbl.Columns["CkDate"]]));
                        sbText.Append("   Federal Program: ");
                        sbText.Append((string)row[tbl.Columns["FedPgmName"]]);
                        sbText.Append("   Amount: ");
                        sbText.Append(row[tbl.Columns["Amount"]].ToString());

                        if (sbHtml.Length != 0)
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
                    if (sbText.Length != 0)
                    {
                        textBody = string.Format(text_body, curStudentFirst, curStudentLast, curSchoolName, sbText.ToString(), curSchoolID);
                        htmlBody = string.Format(html_body, curStudentFirst, curStudentLast, curSchoolName, sbHtml.ToString(), curSchoolID);

                        try
                        {
                            Response r = await SendEmail(apiKey, curStudentFirst, curStudentLast, curStudentEmail, curSchoolContact, curSchoolContactEmail, fromEmail, fromName, textBody, htmlBody);
                            if (r.StatusCode == System.Net.HttpStatusCode.Accepted && paymentList.Count != 0)
                            {
                                UpdatePayments(curStudentFirst, curStudentLast, paymentList, DateTime.Now);
                                UpdateStudent(curStudentFirst, curStudentLast, curStudentID, DateTime.Now);
                                Status = $"Updated payment records & notes for {curStudentFirst} {curStudentLast}";
                            }
                            //else if (r.StatusCode == System.Net.HttpStatusCode.PartialContent)
                            //{
                            //    Status = $"Missing email for {curStudentFirst} {curStudentLast}";
                            //}
                            //else if (r.StatusCode == System.Net.HttpStatusCode.PreconditionFailed)
                            //{
                            //    Status = $"Error occurred trying to send email for {curStudentFirst} {curStudentLast}";
                            //}
                        }
                        catch (Exception ex)
                        {
                            Status = $"Error sending email or updating tables: {ex.Message}";
                            LogIt.LogError(Status);
                        }

                    }
                    Status = $"Processing complete for {school}.";
                    LogIt.LogInfo(Status);
                }
            }
            if (isAutoRun)
                Application.Exit();
        }

        #endregion

        #region methods

        delegate void StringArgReturningVoidDelegate(string status);
        private void ShowStatus(string status)
        {
            // InvokeRequired required compares the thread ID of the  
            // calling thread to the thread ID of the creating thread.  
            // If these threads are different, it returns true.  
            if (txtStatus.InvokeRequired)
            {
                StringArgReturningVoidDelegate d = new StringArgReturningVoidDelegate(ShowStatus);
                this.Invoke(d, new object[] { status });
            }
            else
            {
                txtStatus.Text = status;
            }
        }


        private string GetSetting(XmlDocument doc, string settingName)
        {
            string response = "";
            try
            {
                response = ((XmlElement)doc.SelectSingleNode($"/Settings/setting[@name='{settingName}']")).GetAttribute("value");

            }
            catch (Exception)
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
                using (ODBCClass o = new ODBCClass("MySql"))
                {
                    var procName = "select SchoolID from SchoolInfo";

                    using (OdbcCommand oCommand = o.GetCommand(procName))
                    {
                        using (OdbcDataAdapter oAdapter = new OdbcDataAdapter(oCommand))
                        {
                            oAdapter.Fill(ds);
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                var msg = $"Error getting schools: {ex.Message}";
                LogIt.LogError(msg);
            }
            return ds;
        }

        private void UpdateStudent(string studentFirst, string studentLast, int student, DateTime now)
        {
            string update_sql = $"Update Student set Notes = concat(Notes, '{"\r\n"}', '{now.ToString("MM/dd/yyyy")}: EFT email sent.') where StudentID = ({student});";

            using (ODBCClass o = new ODBCClass("MySql"))
            {
                using (OdbcCommand oCommand = o.GetCommand(update_sql))
                {
                    oCommand.CommandText = update_sql;
                    try
                    {
                        var rows = oCommand.ExecuteNonQuery();
                        LogIt.LogInfo($"Updated notes for {studentFirst} {studentLast} (student ID {student}), Rows affected: {rows}");
                    }
                    catch (Exception ex)
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

            using (ODBCClass o = new ODBCClass("MySql"))
            {
                using (OdbcCommand oCommand = o.GetCommand(update_sql))
                {
                    oCommand.CommandText = update_sql;
                    oCommand.Parameters.AddWithValue("@date", now.ToString("yyyy-MM-dd HH:mm:ss"));
                    oCommand.Parameters.AddWithValue("@pymts", inList);
                    try
                    {
                        var rows = oCommand.ExecuteNonQuery();
                        LogIt.LogInfo($"Updated payments for {studentFirst} {studentLast}, IDs {inList}, Rows affected: {rows}");
                    }
                    catch (Exception ex)
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
                Status = $"Getting payments made for {school}";
                LogIt.LogInfo(Status);

                using (ODBCClass o = new ODBCClass("MySql"))
                {
                    var procName = "payments_made(?,?,?)";
                    var parms = new KeyValuePair<string, string>[]{
                                    new KeyValuePair<string,string>("_school", school),
                                    new KeyValuePair<string,string>("_fromDate", fromDate.ToString("yyyy-MM-dd")),
                                    new KeyValuePair<string,string>("_toDate", toDate.ToString("yyyy-MM-dd"))
                                };

                    using (OdbcCommand oCommand = o.GetCommand(procName, parms))
                    {
                        using (OdbcDataAdapter oAdapter = new OdbcDataAdapter(oCommand))
                        {
                            oAdapter.Fill(ds);
                        }
                    }

                }
            }
            catch (Exception)
            {
            }
            return ds;
        }

        #endregion

        private async Task<Response> SendEmail(string apiKey, string studFirst, string studLast, string emailAddress, string fromEmail, string fromName, string schoolContact, string schoolEmail, string txtBody, string htmBody)
        {
            return await Task<Response>.Run(async () =>
             {
                 Response resp = new Response(System.Net.HttpStatusCode.PartialContent, null, null);
                 var client = new SendGridClient(apiKey);
                 var from = new EmailAddress(fromEmail, fromName);
                 var subject = "Financial Aid Awarded";
                 var to = new EmailAddress(sendTo == "student" ? emailAddress : sendTo, $"{studFirst} {studLast}");
                 var plainTextContent = txtBody;
                 var htmlContent = htmBody;

                 if (to.Email == "")
                 {
                     Status = $"Missing email for {studFirst} {studLast}";
                     LogIt.LogError(Status);
                 }
                 else
                 {
                     try
                     {
                         var msg = MailHelper.CreateSingleEmail(from, to, subject, plainTextContent, htmlContent);

                         // add any CCs from settings
                         if (CCs.Count > 0)
                         {
                             var ccList = new List<EmailAddress>();
                             foreach (var cc in CCs)
                             {
                                 ccList.Add(new EmailAddress(cc, $"{studFirst} {studLast}"));
                             }
                             msg.AddCcs(ccList);
                         }

                         // add BCC from settings if supplied
                         if (bccTo != "")
                             msg.AddBcc(new EmailAddress(bccTo == "school" ? schoolEmail : bccTo, schoolContact));

                         resp = await client.SendEmailAsync(msg);

                         Status = $"Sent email to {to.Name} <{to.Email}>, result={resp.StatusCode.ToString()}";
                         LogIt.LogInfo(Status);

                     }
                     catch (Exception ex)
                     {
                         Status = $"Could not send email: {ex.Message}";
                         LogIt.LogError(Status);
                         resp.StatusCode = System.Net.HttpStatusCode.PreconditionFailed;
                     }
                 }
                 return resp;
             });
        }

    }

}

