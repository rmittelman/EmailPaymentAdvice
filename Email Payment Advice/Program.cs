using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Configuration;
using System.Data.Odbc;
using SendGrid;
using SendGrid.Helpers.Mail;
using Aimm.Logging;

namespace EmailPaymentAdvice
{
    static class Program
    {


        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }







    ////private static Response response;
    ////    private static string text_body = "Dear {0} {1},\n\nYour Financial Aid to attend {2} has been posted "
    ////                + "to your student account as follows:\n\n{3}\n\n"
    ////                + "Please understand that as a student you have the right to cancel all or a portion of a loan "
    ////                + "or loan disbursements.  The request must be in writing and must be received by the school "
    ////                + "within 30 days from the date of the loan disbursement.  If you chose this option, you may "
    ////                + "owe the school money which may be payable at the time of cancellation.\n\n"
    ////                + "If you have any questions, please contact {4}’s Financial Aid Department.\n\n"
    ////                + "Please do not reply to this email because this email box is not monitored.";

    ////    private static string html_body = "<p>Dear {0} {1},</p><p>Your Financial Aid to attend {2} has been posted "
    ////                + "to your student account as follows:</p><p><pre>{3}</pre></p>"
    ////                + "<p>Please understand that as a student you have the right to cancel all or a portion of a loan "
    ////                + "or loan disbursements.  The request must be in writing and must be received by the school "
    ////                + "within 30 days from the date of the loan disbursement.  If you chose this option, you may "
    ////                + "owe the school money which may be payable at the time of cancellation.</p>"
    ////                + "<p>If you have any questions, please contact {4}’s Financial Aid Department.</p>"
    ////                + "<p>Please do not reply to this email because this email box is not monitored.</p>";

    ////    private static string textBody;
    ////    private static string htmlBody;
    ////    private static string curStudentFirst;
    ////    private static string curStudentLast;
    ////    private static string curSchoolName;
    ////    private static string curSchoolID;
    ////    private static string curSchoolContact;
    ////    private static string curSchoolContactEmail;
    ////    private static string curStudentEmail;
    ////    private static int fromDaysAgo = -1;
    ////    private static int toDaysAgo = -1;
    ////    private static string sendTo;
    ////    private static string bccTo;
    ////    private static List<string> CCs = new List<string>();
    ////    private static List<string> Schools = new List<string>();
    ////    private static string apiKey;

    //    //private static string update_sql = "Update Payments p set p.AdviceDate = '{0}' where p.PaymentID in({1});";

    //    static void Main(string[] args)
    //    {
    ////        LogIt.LogMethod();
    ////        DateTime tempFrom;
    ////        DateTime tempTo;
    ////        DateTime fromDate;
    ////        DateTime toDate;

    ////        sendTo = ConfigurationManager.AppSettings["SendTo"];
    ////        bccTo = ConfigurationManager.AppSettings["SendBCC"];
    ////        var tempSchools = ConfigurationManager.AppSettings["Schools"];
    ////        var tempCCs = ConfigurationManager.AppSettings["SendCC"];
    ////        fromDaysAgo = Convert.ToInt32(ConfigurationManager.AppSettings["FromDaysAgo"]);
    ////        toDaysAgo = Convert.ToInt32(ConfigurationManager.AppSettings["ToDaysAgo"]);
    ////        apiKey = ConfigurationManager.AppSettings["SENDGRID_API_KEY"];
    ////        if (tempSchools.Length == 0 || fromDaysAgo < 0 || toDaysAgo < 0) // !DateTime.TryParse(args[0], out temp) || !DateTime.TryParse(args[1], out temp))
    ////        {
    ////            LogIt.LogError($"Bad data supplied: Schools={tempSchools}, FromDaysAgo={fromDaysAgo}, ToDaysAgo={toDaysAgo}");
    ////        }
    ////        else
    ////        {
    ////            // if valid from & to dates supplied in command line, use those instead
    ////            if(args.Length == 2)
    ////            {
    ////                try
    ////                {
    ////                    if (DateTime.TryParse(args[0], out tempFrom) && DateTime.TryParse(args[1], out tempTo))
    ////                    {
    ////                        fromDate = tempFrom;
    ////                        toDate = tempTo;
    ////                        LogIt.LogInfo($"Got beginning and ending dates from command-line arguments: FromDate={fromDate}, ToDate={toDate}");
    ////                    }
    ////                    else
    ////                    {
    ////                        fromDate = DateTime.Today.AddDays(-fromDaysAgo);
    ////                        toDate = DateTime.Today.AddDays(-toDaysAgo).AddHours(23).AddMinutes(59).AddSeconds(59);
    ////                        LogIt.LogInfo($"Got beginning and ending dates from config file: FromDate={fromDate}, ToDate={toDate}");
    ////                    }

    ////                }
    ////                catch (Exception)
    ////                {
    ////                    fromDate = DateTime.Today.AddDays(-fromDaysAgo);
    ////                    toDate = DateTime.Today.AddDays(-toDaysAgo).AddHours(23).AddMinutes(59).AddSeconds(59);
    ////                    LogIt.LogInfo($"Error in command-line dates, got beginning and ending dates from config file: FromDate={fromDate}, ToDate={toDate}");
    ////                }

    ////            }
    //            else
    //            {
    ////                fromDate = DateTime.Today.AddDays(-fromDaysAgo);
    ////                toDate = DateTime.Today.AddDays(-toDaysAgo).AddHours(23).AddMinutes(59).AddSeconds(59);
    ////                LogIt.LogInfo($"Got beginning and ending dates from config file: FromDate={fromDate}, ToDate={toDate}");
    //            }


    ////            Schools = tempSchools.Split(',').ToList<string>();
    ////            CCs = tempCCs.Split(',').ToList<string>();

    ////            // loop thru schools, process each
    ////            foreach (var school in Schools)
    ////            {
    ////                StringBuilder sbText = new StringBuilder();
    ////                StringBuilder sbHtml = new StringBuilder();
    ////                List<int> paymentList = new List<int>();

    ////                LogIt.LogInfo($"Processing: School={school}, From={fromDate}, To={toDate}");
    ////                DataSet ds = new DataSet();
    ////                try
    ////                {
    ////                    LogIt.LogInfo("Getting payments made.");
    ////                    using (ODBCClass o = new ODBCClass("MySql"))
    ////                    {
    ////                        var procName = "payments_made(?,?,?)";
    ////                        var parms = new KeyValuePair<string, string>[]{
    ////                    new KeyValuePair<string,string>("_school", school),
    ////                    new KeyValuePair<string,string>("_fromDate", fromDate.ToString("yyyy-MM-dd")),
    ////                    new KeyValuePair<string,string>("_toDate", toDate.ToString("yyyy-MM-dd"))
    ////                };

    ////                        using (OdbcCommand oCommand = o.GetCommand(procName, parms))
    ////                        {
    ////                            using (OdbcDataAdapter oAdapter = new OdbcDataAdapter(oCommand))
    ////                            {
    ////                                oAdapter.Fill(ds);
    ////                            }
    ////                        }
    ////                    }

    ////                }
    ////                catch (Exception ex)
    ////                {
    ////                    LogIt.LogError($"Error getting payments: {ex.Message}");
    ////                    //if (System.Windows.Forms.Application.MessageLoop)
    ////                    //    System.Windows.Forms.Application.Exit();
    ////                    //else
    ////                    Environment.Exit(1);
    ////                }

    ////                if (ds.Tables[0].Rows.Count == 0)
    ////                {
    ////                    LogIt.LogInfo($"No payments to process for {school} between {fromDate} and {toDate}");
    ////                    Environment.Exit(1);
    ////                }

    //                // loop thru table rows and build emails
    //                DataTable tbl = ds.Tables[0];
    //                int curStudentID = 0;
    //                int curSchool_ID = 0;

    //                DataColumn idCol = tbl.Columns["StudentID"];

    ////                foreach (DataRow row in tbl.Rows)
    ////                {
    ////                    // when we encounter a new student, process the email for prior student
    ////                    if ((int)row[idCol] != curStudentID)
    ////                    {
    ////                        // we've hit a new student, so if not the very first student, send email
    ////                        if (curStudentID != 0)
    ////                        {
    ////                            textBody = string.Format(text_body, curStudentFirst, curStudentLast, curSchoolName, sbText.ToString(), curSchoolID);
    ////                            htmlBody = string.Format(html_body, curStudentFirst, curStudentLast, curSchoolName, sbHtml.ToString(), curSchoolID);

    ////                            try
    ////                            {
    ////                                // send the email, update payments and student if successful
    ////                                Execute(apiKey, curStudentFirst, curStudentLast, curStudentEmail, curSchoolContact, curSchoolContactEmail, textBody, htmlBody).Wait();
    ////                                if (response.StatusCode == System.Net.HttpStatusCode.Accepted && paymentList.Count != 0)
    ////                                {
    ////                                    update_payments(paymentList, DateTime.Now);
    ////                                    update_student(curStudentID, DateTime.Now);
    ////                                }
    ////                            }
    ////                            catch (Exception ex)
    ////                            {
    ////                                LogIt.LogError($"Error sending email: {ex.Message}");
    ////                            }

    ////                        }

    ////                        // set current values, clear list of payments
    ////                        curStudentID = (int)row[tbl.Columns["StudentID"]];
    ////                        curSchool_ID = (int)row[tbl.Columns["School_ID"]];
    ////                        curSchoolID = (string)row[tbl.Columns["SchoolID"]] ?? "";
    ////                        curSchoolName = (string)row[tbl.Columns["SchoolName"]] ?? "";
    ////                        curSchoolContact = (string)row[tbl.Columns["ContactName"]] ?? "";
    ////                        curSchoolContactEmail = (string)row[tbl.Columns["ContactEmail"]] ?? "";
    ////                        curStudentFirst = (string)row[tbl.Columns["FirstName"]] ?? "";
    ////                        curStudentLast = (string)row[tbl.Columns["LastName"]] ?? "";
    ////                        curStudentEmail = (string)row[tbl.Columns["Email"]] ?? "";

    ////                        sbText = new StringBuilder();
    ////                        sbHtml = new StringBuilder();
    ////                        paymentList = new List<int>();

    ////                    }

    ////                    // append to list of payments and stringBuilders
    ////                    paymentList.Add((int)row[tbl.Columns["PaymentID"]]);

    ////                    if (sbText.Length != 0)
    ////                        sbText.Append("\n");
    ////                    sbText.Append("Check #: ");
    ////                    sbText.Append(row[tbl.Columns["CkNo"]].ToString());
    ////                    sbText.Append("   Date: ");
    ////                    sbText.Append(string.Format("{0:M/d/yyyy}", row[tbl.Columns["CkDate"]]));
    ////                    sbText.Append("   Federal Program: ");
    ////                    sbText.Append((string)row[tbl.Columns["FedPgmName"]]);
    ////                    sbText.Append("   Amount: ");
    ////                    sbText.Append(row[tbl.Columns["Amount"]].ToString());

    ////                    if (sbHtml.Length != 0)
    ////                        sbHtml.Append("\n");
    ////                    sbHtml.Append("Check #: ");
    ////                    sbHtml.Append(row[tbl.Columns["CkNo"]].ToString());
    ////                    sbHtml.Append("   Date: ");
    ////                    sbHtml.Append(string.Format("{0:M/d/yyyy}", row[tbl.Columns["CkDate"]]));
    ////                    sbHtml.Append("   Federal Program: ");
    ////                    sbHtml.Append((string)row[tbl.Columns["FedPgmName"]]);
    ////                    sbHtml.Append("   Amount: ");
    ////                    sbHtml.Append(row[tbl.Columns["Amount"]].ToString());

    ////                }

    ////                // send email and update if any payment details accumulated for last student in school.
    ////                if (sbText.Length != 0)
    ////                {
    ////                    textBody = string.Format(text_body, curStudentFirst, curStudentLast, curSchoolName, sbText.ToString(), curSchoolID);
    ////                    htmlBody = string.Format(html_body, curStudentFirst, curStudentLast, curSchoolName, sbHtml.ToString(), curSchoolID);

    ////                    try
    ////                    {
    ////                        Execute(apiKey, curStudentFirst, curStudentLast, curStudentEmail, curSchoolContact, curSchoolContactEmail, textBody, htmlBody).Wait();

    ////                        // if successful, update any payments
    ////                        if (response.StatusCode == System.Net.HttpStatusCode.Accepted && paymentList.Count != 0)
    ////                        {
    ////                            update_payments(paymentList, DateTime.Now);
    ////                            update_student(curStudentID, DateTime.Now);
    ////                        }
    ////                    }
    ////                    catch (Exception ex)
    ////                    {

    ////                        System.Diagnostics.Debug.WriteLine("Error: " + ex.Message);
    ////                    }

    ////                }
    ////            }

    //            LogIt.LogInfo("Program execution complete");
    //        }

    //    }

    ////    private static void update_student(int student, DateTime now)
    ////    {
    ////        string update_sql = $"Update Student set Notes = concat(Notes, '{"\r\n"}', '{now.ToString("MM/dd/yyyy")}: EFC email sent.') where StudentID = ({student});";
    ////        //            update fc5.Student set Notes = concat(Notes, '\n', '09/27/2017: EFC email sent.')
    ////        //where StudentID in (1);

    ////        using (ODBCClass o = new ODBCClass("MySql"))
    ////        {
    ////            using (OdbcCommand oCommand = o.GetCommand(update_sql))
    ////            {
    ////                oCommand.CommandText = update_sql;
    ////                //oCommand.Parameters.AddWithValue("@date", now.ToString("yyyy-MM-dd HH:mm:ss"));
    ////                //oCommand.Parameters.AddWithValue("@pymts", inList);
    ////                try
    ////                {
    ////                    var rows = oCommand.ExecuteNonQuery();
    ////                    LogIt.LogInfo($"Updated student ID {student}, Rows affected: {rows}");
    ////                }
    ////                catch (Exception ex)
    ////                {
    ////                    LogIt.LogError($"Could not update student {student}: {ex.Message}");
    ////                }
    ////            }
    ////        }
    ////    }

    ////    private static void update_payments(List<int> paymentList, DateTime now)
    ////    {
    ////        string inList = string.Join(",", paymentList.ConvertAll(Convert.ToString));
    ////        string update_sql = $"Update Payments p set p.AdviceDate = '{now.ToString("yyyy-MM-dd HH:mm:ss")}' where p.PaymentID in({inList});";

    ////        using (ODBCClass o = new ODBCClass("MySql"))
    ////        {
    ////            using (OdbcCommand oCommand = o.GetCommand(update_sql))
    ////            {
    ////                oCommand.CommandText = update_sql;
    ////                oCommand.Parameters.AddWithValue("@date", now.ToString("yyyy-MM-dd HH:mm:ss"));
    ////                oCommand.Parameters.AddWithValue("@pymts", inList);
    ////                try
    ////                {
    ////                    var rows = oCommand.ExecuteNonQuery();
    ////                    LogIt.LogInfo($"Updated payment IDs {inList}, Rows affected: {rows}");
    ////                }
    ////                catch (Exception ex)
    ////                {
    ////                    LogIt.LogError($"Could not update payments: {ex.Message}");
    ////                }
    ////            }
    ////        }
    ////    }

    //    static async Task Execute(string apiKey, string studFirst, string studLast, string emailAddress, string schoolContact, string schoolEmail, string txtBody, string htmBody)
    //    {
    //        try
    //        {
    //            //var apiKey = Environment.GetEnvironmentVariable("SENDGRID_API_KEY");
    //            var client = new SendGridClient(apiKey);
    //            var from = new EmailAddress("info@ShamrocksFA.com", "Shamrocks Unlimited");
    //            var subject = "Financial Aid Awarded";
    //            var to = new EmailAddress(sendTo == "student" ? emailAddress : sendTo, $"{studFirst} {studLast}");
    //            var plainTextContent = txtBody;
    //            var htmlContent = htmBody;

    //            if(to.Email == "")
    //            {
    //                LogIt.LogError($"Missing email for {studFirst} {studLast}");
    //                response = null;
    //            }
    //            else
    //            {
    //                var msg = MailHelper.CreateSingleEmail(from, to, subject, plainTextContent, htmlContent);

    //                // add any CCs from settings
    //                if (CCs.Count > 0)
    //                {
    //                    var ccList = new List<EmailAddress>();
    //                    foreach (var cc in CCs)
    //                    {
    //                        ccList.Add(new EmailAddress(cc, $"{studFirst} {studLast}"));
    //                    }
    //                    msg.AddCcs(ccList);
    //                }

    //                // add BCC from settings if supplied
    //                if (bccTo != "")
    //                    msg.AddBcc(new EmailAddress(bccTo == "school" ? schoolEmail : bccTo, schoolContact));

    //                response = await client.SendEmailAsync(msg);
    //                LogIt.LogInfo($"Sent email to {to.Name} <{to.Email}>, result={response.StatusCode.ToString()}");
    //            }

    //        }
    //        catch (Exception ex)
    //        {
    //            LogIt.LogError($"Could not send email: {ex.Message}");
    //        }
    //    }

    //}
}
