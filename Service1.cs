using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
//using System.Threading;
//using System.Threading.Tasks;
using System.Timers;

namespace Email_Send_WinService
{
    public partial class Service1 : ServiceBase
    {
        private Timer timer1 = null;
        int PickUpDays = 0;
        int ReminderDays = 0;
        int SPRunMinutes = 0;
        // int ServiceRunTime = 0;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {

            try
            {
                LogService.WriteErrorLog("Email send window service started");

                int runTime = Convert.ToInt32(ConfigurationManager.AppSettings["SendMailRuntimeInMin"]);
                double milliSeconds = TimeSpan.FromMinutes(runTime).TotalMilliseconds;
                timer1 = new Timer();
                this.timer1.Interval = milliSeconds;
                this.timer1.Elapsed += new System.Timers.ElapsedEventHandler(this.timer1_Tick);
                timer1.Enabled = true;



            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(ex);
            }
        }

        protected override void OnStop()
        {
            timer1.Enabled = false;
            LogService.WriteErrorLog("Email send window service stopped");
        }

        private void timer1_Tick(object sender, ElapsedEventArgs e)
        {
            //LogService.WriteErrorLog("timer1_Tick");           
            //LogService.WriteErrorLog("GetBadgeReportSettings completed");
            GetBadgeReportSettings();
            Send_Email();
            //Method for security clearance
            SendReminderMail();
            //
            SaveSecurityClearanceData();
            //Reminder audit mail for authsigner
            SendRemindeAuditMail();
            //Citation mail, over due for authsigner
            SendRemindeNovMail();
            //send email for missed audit on yesterday..(once in a day)
            TrySendMissingAuditNoticemail();
        }

        private void Send_Email()
        {
            try
            {
                // LogService.WriteErrorLog("Send_Email");
                DAL dal = new DAL();
                string Dkey = "($h@r!(u!8*MW4oB1VmL5GIwBIjqFYQntHT0CMi2uEYAmBkwxpvsbLQ6KX1SCno9XQ==";
                //EncryptDecryptPassword encypt = new EncryptDecryptPassword();
                List<EmailConfigModel> emailConfigs = dal.GetEmailConfigs(string.Empty);
                List<EmailDataModel> emailDatas = dal.GetEmailData(string.Empty);
                foreach (EmailDataModel data in emailDatas)
                {
                    //string EncryptedAppName = EncryptDecryptPassword.EncryptText(data.ApplicationName.Trim(), Dkey);
                    //EmailConfigModel config = emailConfigs.Where(x => x.ApplicationName == EncryptedAppName).FirstOrDefault();
                    EmailConfigModel config = emailConfigs.Where(x => x.ApplicationName == data.ApplicationName).FirstOrDefault();
                    if (config != null && !string.IsNullOrEmpty(config.ApplicationName))
                    {
                        try
                        {
                            using (MailMessage mail = new MailMessage())
                            {
                                string DecryptedFromEmail = EncryptDecryptPassword.DecryptText(config.FromMail, Dkey);
                                string DecryptedEmail = EncryptDecryptPassword.DecryptText(config.Username, Dkey);
                                string DecryptedPassword = EncryptDecryptPassword.DecryptText(config.Password, Dkey);
                                string DecryptedHost = EncryptDecryptPassword.DecryptText(config.SMTPHost, Dkey);
                                string DecryptedPort = EncryptDecryptPassword.DecryptText(config.Port, Dkey);

                                mail.From = new MailAddress(DecryptedFromEmail);
                                mail.To.Add(data.ToMail);
                                if (!string.IsNullOrEmpty(data.CCMail))
                                    mail.CC.Add(data.CCMail);
                                if (!string.IsNullOrEmpty(data.BCCMail))
                                    mail.Bcc.Add(data.BCCMail);
                                mail.Subject = data.Subject;
                                mail.Body = data.MailBody;
                                mail.IsBodyHtml = true;
                                string attachmentPath = data.AttachmentPath;
                                if (!string.IsNullOrEmpty(attachmentPath) && System.IO.File.Exists(attachmentPath))
                                {
                                    mail.Attachments.Add(new Attachment(attachmentPath));
                                }
                                //Update Row of email send with IsMailSent is 2, shows it is Processing row and if Send(mail) get failed status remainsame or it get sent status set to 1
                                dal.UpdateRowProcessStatus(data.Id);
                                using (SmtpClient smtp = new SmtpClient(DecryptedHost, Convert.ToInt32(DecryptedPort)))
                                {
                                    smtp.Credentials = new NetworkCredential(DecryptedEmail, DecryptedPassword);
                                    smtp.EnableSsl = true;                                    
                                    smtp.Send(mail);
                                }
                            }

                            dal.UpdateMailSentStatus(data.Id);
                        }
                        catch (Exception ex)
                        {
                            LogService.WriteErrorLog(DateTime.Now + string.Format(" : Mail not sent to Email - {0}, subject - {1}. Error - {2} ", data.ToMail, data.Subject, ex.Message));
                        }
                    }
                    else
                    {
                        LogService.WriteErrorLog(DateTime.Now + string.Format(" : Mail not sent to Email - {0}, subject - {1} because email configuration details not available for application - {2} ", data.ToMail, data.Subject, data.ApplicationName));
                    }
                }

            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(DateTime.Now + string.Format(" : Error Found in the Send_Email(). Exception - {0} ", ex.Message));
            }
        }

        private void SaveSecurityClearanceData()
        {
            try
            {
                //LogService.WriteErrorLog("SaveSecurityClearanceData");
                DAL_SVMS dal = new DAL_SVMS();

                dal.SaveSecurityClearanceData(SPRunMinutes);
            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(DateTime.Now + string.Format(" : Error Found in the SaveSecurityClearanceData(). Exception - {0} ", ex.Message));

            }

        }
        private void GetBadgeReportSettings()
        {
            try
            {
                //LogService.WriteErrorLog("GetBadgeReportSettings");
                DAL_SVMS dal = new DAL_SVMS();
                DataTable badgeReport = dal.GetReminderClearanceData("BR");
                if (badgeReport != null)
                {
                    if (badgeReport.Rows.Count > 0)
                    {

                        PickUpDays = Convert.ToInt16(badgeReport.Rows[0]["PickUpDays"]);
                        ReminderDays = Convert.ToInt16(badgeReport.Rows[0]["ReminderDays"]);
                        SPRunMinutes = Convert.ToInt16(badgeReport.Rows[0]["SPRunMinutes"]);

                    }
                }

            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(DateTime.Now + string.Format(" : Error Found in the GetBadgeReportSettings(). Exception - {0} ", ex.Message));
            }

        }
        public void SendReminderMail()
        {
            try
            {
                //LogService.WriteErrorLog("SendReminderMail");
                DAL_SVMS dal = new DAL_SVMS();
                //Get list of records, due more than 15 days 
                DataTable dt = dal.GetReminderClearanceData("RD");
                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        DataTable badgeReportEmail = dal.GetReminderClearanceData("BRE");
                        string CC = "";
                        string BCC = "";

                        for (int e = 0; e < badgeReportEmail.Rows.Count; e++)
                        {
                            string Type = badgeReportEmail.Rows[e]["Type"].ToString();
                            if (Type == "CC")
                            {
                                CC += badgeReportEmail.Rows[e]["Email"].ToString() + ",";

                            }
                            else
                            {
                                BCC += badgeReportEmail.Rows[e]["Email"].ToString() + ",";
                            }
                        }

                        CC = (CC != "" ? CC.Remove(CC.Length - 1, 1) : "");
                        BCC = (BCC != "" ? BCC.Remove(BCC.Length - 1, 1) : "");
                        List<ReminderClearanceData> listName = dt.AsEnumerable().Select(m => new ReminderClearanceData()
                        {
                            Id = m.Field<int>("Id"),
                            AuthSignerFirstName = m.Field<string>("AuthSignerFirstName"),
                            AuthSignerLastName = m.Field<string>("AuthSignerLastName"),
                            UserFirstName = m.Field<string>("UserFirstName"),
                            UserLastName = m.Field<string>("UserLastName"),
                            ClearanceDate = m.Field<DateTime?>("ClearanceDate"),
                            NotificationDate = m.Field<DateTime?>("NotificationDate"),
                            Email = m.Field<string>("Email"),
                            CompanyId = m.Field<int>("CompanyId")
                        }).ToList();

                        List<int> compIds = listName.Select(x => x.CompanyId).Distinct().ToList();
                        foreach (int compId in compIds)
                        {
                            List<ReminderClearanceData> companyWiseData = listName.Where(x => x.CompanyId == compId).ToList();

                            string authSignerEmail = string.Join(",", companyWiseData.Select(x => x.Email).Distinct().ToArray());
                            var distinctData = companyWiseData.Select(x => new { x.NotificationDate, x.ClearanceDate, x.UserFirstName, x.UserLastName, x.Id }).Distinct().ToList();
                            foreach (var item in distinctData)
                            {
                                string subject = "Badge Security Clearance Reminder";
                                string htmlBody = string.Empty;
                                string AssemblyPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location).ToString();
                                using (StreamReader sr = new StreamReader(AssemblyPath + "/EmailTemplate/ReminderMail.html"))
                                {
                                    htmlBody = sr.ReadToEnd();
                                }
                                string callbackUrl = ConfigurationManager.AppSettings["SVMSGUILink"];
                                // DateTime NotificationDate = Convert.ToDateTime(dt.Rows[i]["NotificationDate"]).AddDays(PickUpDays);
                                DateTime? NotificationDate = Convert.ToDateTime(item.NotificationDate);
                                if (NotificationDate.HasValue)
                                {
                                    string formattedNotificationDate = NotificationDate?.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                                    LogService.WriteErrorLog($"Notification Date is: {formattedNotificationDate}");
                                    htmlBody = htmlBody.Replace("#NotificationDate", formattedNotificationDate);
                                }
                                else
                                {
                                    htmlBody = htmlBody.Replace("#NotificationDate", "");

                                }
                                DateTime? ClearanceDate = Convert.ToDateTime(item.ClearanceDate);
                                if (ClearanceDate.HasValue)
                                {
                                    string formattedClearanceDate = ClearanceDate?.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                                    LogService.WriteErrorLog($"Clearance Date is: {formattedClearanceDate}");
                                    htmlBody = htmlBody.Replace("#ClearanceDate", formattedClearanceDate);
                                }
                                else
                                {
                                    htmlBody = htmlBody.Replace("#ClearanceDate", "");

                                }
                                //htmlBody = htmlBody.Replace("#NotificationDate", NotificationDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture));
                                //htmlBody = htmlBody.Replace("#ClearanceDate", ClearanceDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture));
                                htmlBody = htmlBody.Replace("#FirstName", item.UserFirstName.ToString());
                                htmlBody = htmlBody.Replace("#LastName", item.UserLastName.ToString());
                                htmlBody = htmlBody.Replace("hrefCode", callbackUrl);
                                htmlBody = htmlBody.Replace("#PickUpDays", PickUpDays.ToString());

                                DAL dal_email = new DAL();
                                if (!string.IsNullOrEmpty(authSignerEmail))
                                {
                                    dal_email.SendEmailUsingService("BadgeReport", authSignerEmail, CC, BCC, subject, htmlBody, "");

                                }
                                //  SendMail_SVMSReminder(dt.Rows[i]["Email"].ToString(), "", "", subject, htmlBody, Convert.ToInt32(dt.Rows[i]["Id"]));

                                dal = new DAL_SVMS();
                                dal.UpdateMailSentStatus(Convert.ToInt32(item.Id));
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(DateTime.Now + string.Format(" : Error Found in the SendReminderMail(). Exception - {0} ", ex.Message));

            }



        }



        public void SendRemindeAuditMail()
        {
            try
            {
                //LogService.WriteErrorLog("SendReminderMail");
                DAL_SVMS dal = new DAL_SVMS();
                DataTable dt = dal.GetReminderAuditData("RD");
                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {

                        DataTable badgeReportEmail = dal.GetReminderAuditData("BRE");
                        string CC = "";
                        string BCC = "";

                        for (int e = 0; e < badgeReportEmail.Rows.Count; e++)
                        {
                            string Type = badgeReportEmail.Rows[e]["Type"].ToString();
                            if (Type == "CC")
                            {
                                CC += badgeReportEmail.Rows[e]["Email"].ToString() + ",";

                            }
                            else
                            {
                                BCC += badgeReportEmail.Rows[e]["Email"].ToString() + ",";
                            }
                        }

                        CC = (CC != "" ? CC.Remove(CC.Length - 1, 1) : "");
                        BCC = (BCC != "" ? BCC.Remove(BCC.Length - 1, 1) : "");

                        List<ReminderAuditData> listName = dt.AsEnumerable().Select(m => new ReminderAuditData()
                        {
                            Id = m.Field<int>("Id"),
                            AuthSignerFirstName = m.Field<string>("AuthSignerFirstName"),
                            AuthSignerLastName = m.Field<string>("AuthSignerLastName"),
                            //UserFirstName = m.Field<string>("UserFirstName"),
                            //UserLastName = m.Field<string>("UserLastName"),
                            //ClearanceDate = m.Field<DateTime>("ClearanceDate"),
                            AuditToDate = m.Field<DateTime>("AuditToDate"),
                            AuditFromDate = m.Field<DateTime>("AuditFromDate"),
                            AuditName = m.Field<string>("AuditName"),
                            Email = m.Field<string>("Email"),
                            CompanyId = m.Field<int>("CompanyId")
                        }).ToList();

                        List<int> compIds = listName.Select(x => x.CompanyId).Distinct().ToList();
                        foreach (int compId in compIds)
                        {
                            List<ReminderAuditData> companyWiseData = listName.Where(x => x.CompanyId == compId).ToList();
                            string authSignerEmail = string.Join(",", companyWiseData.Select(x => x.Email).ToArray());
                            var distinctData = companyWiseData.Select(x => new { x.AuditToDate, x.Id, x.AuditFromDate, x.AuditName }).Distinct().ToList();

                            foreach (var item in distinctData)
                            {
                                string subject = "Badge Audit Reminder";
                                string htmlBody = string.Empty;
                                string AssemblyPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location).ToString();
                                using (StreamReader sr = new StreamReader(AssemblyPath + "/EmailTemplate/ReminderAuditMail.html"))
                                {
                                    htmlBody = sr.ReadToEnd();
                                }
                                string callbackUrl = ConfigurationManager.AppSettings["SVMSGUILink"];
                                // DateTime NotificationDate = Convert.ToDateTime(dt.Rows[i]["NotificationDate"]).AddDays(PickUpDays);
                                DateTime AuditToDate = Convert.ToDateTime(item.AuditToDate);
                                DateTime AuditFromDate = Convert.ToDateTime(item.AuditFromDate);
                                htmlBody = htmlBody.Replace("#AuditToDate", AuditToDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture));
                                htmlBody = htmlBody.Replace("#AuditFromDate", AuditFromDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture));
                                htmlBody = htmlBody.Replace("#AuditName", item.AuditName.ToString());
                                //htmlBody = htmlBody.Replace("#LastName", item.UserLastName.ToString());
                                htmlBody = htmlBody.Replace("hrefCode", callbackUrl);
                                // htmlBody = htmlBody.Replace("#PickUpDays", PickUpDays.ToString());

                                DAL dal_email = new DAL();
                                if (!string.IsNullOrEmpty(authSignerEmail))
                                {
                                    dal_email.SendEmailUsingService("BadgeAudit", authSignerEmail, CC, BCC, subject, htmlBody, "");

                                }
                                //  SendMail_SVMSReminder(dt.Rows[i]["Email"].ToString(), "", "", subject, htmlBody, Convert.ToInt32(dt.Rows[i]["Id"]));

                                dal = new DAL_SVMS();
                                dal.UpdateAuditMailSentStatus(Convert.ToInt32(item.Id));
                            }
                        }

                    }
                }

            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(DateTime.Now + string.Format(" : Error Found in the SendReminderAuditMail(). Exception - {0} ", ex.Message));

            }



        }

        public void SendRemindeNovMail()
        {
            try
            {
                //LogService.WriteErrorLog("SendReminderMail");
                DAL_SVMS dal = new DAL_SVMS();
                DataTable dt = dal.GetReminderNovData("RD");
                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        //DataTable badgeReportEmail = dal.GetReminderNovData("BRE");
                        string CC = "";
                        string BCC = "";

                        //for (int e = 0; e < badgeReportEmail.Rows.Count; e++)
                        //{
                        //    string Type = badgeReportEmail.Rows[e]["Type"].ToString();
                        //    if (Type == "CC")
                        //    {
                        //        CC += badgeReportEmail.Rows[e]["Email"].ToString() + ",";

                        //    }
                        //    else
                        //    {
                        //        BCC += badgeReportEmail.Rows[e]["Email"].ToString() + ",";
                        //    }
                        //}

                        //CC = (CC != "" ? CC.Remove(CC.Length - 1, 1) : "");
                        //BCC = (BCC != "" ? BCC.Remove(BCC.Length - 1, 1) : "");

                        List<ReminderNovData> listName = dt.AsEnumerable().Select(m => new ReminderNovData()
                        {
                            CitationId = m.Field<int>("CitationId"),
                            AuthSignerFirstName = m.Field<string>("AuthSignerFirstName"),
                            AuthSignerLastName = m.Field<string>("AuthSignerLastName"),
                            //UserFirstName = m.Field<string>("UserFirstName"),
                            //UserLastName = m.Field<string>("UserLastName"),
                            //ClearanceDate = m.Field<DateTime>("ClearanceDate"),
                            CompanyName = m.Field<string>("CompanyName"),
                            ViolatorFirstName = m.Field<string>("ViolatorFirstName"),
                            ViolatorLastName = m.Field<string>("ViolatorLastName"),
                            Email = m.Field<string>("Email"),
                            CompanyId = m.Field<int>("CompanyId"),
                            NovNo = m.Field<int>("NovNo"),
                            RemedialTrainingAssignedDate = m.Field<DateTime>("RemedialTrainingAssignedDate")
                        }).ToList();

                        List<int> compIds = listName.Select(x => x.CompanyId).Distinct().ToList();

                        foreach (int compId in compIds)
                        {
                            List<ReminderNovData> companyWiseData = listName.Where(x => x.CompanyId == compId).Distinct().ToList();
                            List<string> authsignerformaillst = companyWiseData.Select(x => x.Email).Distinct().ToList();
                            string authSignerEmail = string.Join(",", authsignerformaillst.ToArray());
                            var distinctData = companyWiseData.Select(x => new { x.CitationId, x.NovNo, x.ViolatorFirstName, x.ViolatorLastName, x.RemedialTrainingAssignedDate }).Distinct().ToList();
                            LogService.WriteErrorLog("Email sending to authsigners " + authSignerEmail);
                            foreach (var item in distinctData)
                            {
                                string subject = "Citation OverDue Reminder";
                                string htmlBody = string.Empty;
                                string AssemblyPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location).ToString();
                                using (StreamReader sr = new StreamReader(AssemblyPath + "/EmailTemplate/ReminderCitationMail.html"))
                                {
                                    htmlBody = sr.ReadToEnd();
                                }
                                string callbackUrl = ConfigurationManager.AppSettings["SVMSGUILink"];
                                // DateTime NotificationDate = Convert.ToDateTime(dt.Rows[i]["NotificationDate"]).AddDays(PickUpDays);
                                DateTime RemedialTrainingAssignedDate = Convert.ToDateTime(item.RemedialTrainingAssignedDate);

                                htmlBody = htmlBody.Replace("#RemedialTrainingAssignedDate", RemedialTrainingAssignedDate.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture));
                                htmlBody = htmlBody.Replace("#ViolatorFirstName", item.ViolatorFirstName.ToString());
                                htmlBody = htmlBody.Replace("#ViolatorLastName", item.ViolatorLastName.ToString());
                                htmlBody = htmlBody.Replace("#NovNo", item.NovNo.ToString());
                                htmlBody = htmlBody.Replace("hrefCode", callbackUrl);


                                DAL dal_email = new DAL();
                                if (!string.IsNullOrEmpty(authSignerEmail))
                                {
                                    dal_email.SendEmailUsingService("OverdueCitation", authSignerEmail, CC, BCC, subject, htmlBody, "");

                                }
                                //  SendMail_SVMSReminder(dt.Rows[i]["Email"].ToString(), "", "", subject, htmlBody, Convert.ToInt32(dt.Rows[i]["Id"]));
                                LogService.WriteErrorLog("Email send to authsigners " + authSignerEmail);

                                dal = new DAL_SVMS();
                                dal.UpdateNovMailSentStatus(Convert.ToInt32(item.CitationId));
                            }
                        }


                    }
                }

            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(DateTime.Now + string.Format(" : Error Found in the SendRemindeNovMail(). Exception - {0} ", ex.Message));

            }



        }

        
        public void TrySendMissingAuditNoticemail()
        {
            try
            {
                DAL_SVMS dal = new DAL_SVMS();

                // Check if today's email already sent
                if (!dal.IsProhibitedMissingAuditEmailSentToday())
                {
                    SendMissingAuditNoticemail();
                }
                else
                {
                    LogService.WriteErrorLog("ProhibitedMissingAudit email already sent today. Skipping.");
                }
            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog("Error in TrySendMissingAuditNoticemail(): " + ex.Message);
            }
        }

        public void SendMissingAuditNoticemail()
        {
            try
            {
                DateTime now = DateTime.Now;
                if (now.Hour < 5 || (now.Hour == 5 && now.Minute < 10))
                {
                    LogService.WriteErrorLog($"Skipped sending missing audit mail at {now:yyyy-MM-dd HH:mm:ss}. Waiting until after 3:10 AM.");
                    return;
                }

                // Target previous day’s audit
                DateTime targetDate = now.Date.AddDays(-1);


                DAL_SVMS dal = new DAL_SVMS();
                DataTable dt = dal.GetCompaniesWithMissedAuditsData("MA");

                if (dt != null && dt.Rows.Count > 0)
                {
                    string CC = "";
                    string BCC = "";

                    DataTable emailTable = dal.GetCompaniesWithMissedAuditsData("EMAIL");

                    // Build comma-separated email list
                    string toEmail = "";
                    if (emailTable != null && emailTable.Rows.Count > 0)
                    {
                        List<string> emailList = new List<string>();
                        foreach (DataRow row in emailTable.Rows)
                        {
                            if (row["EmailId"] != DBNull.Value)
                                emailList.Add(row["EmailId"].ToString());
                        }
                        toEmail = string.Join(",", emailList);
                    }

                    // Build full email content listing all missed audits
                    if (!string.IsNullOrEmpty(toEmail))
                    {
                        string subject = "Prohibited Missing Audit";
                        string htmlBody = string.Empty;

                        // Load template
                        string assemblyPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                        using (StreamReader sr = new StreamReader(Path.Combine(assemblyPath, "EmailTemplate", "MissedProhibitedAudit.html")))
                        {
                            htmlBody = sr.ReadToEnd();
                        }

                        string callbackUrl = ConfigurationManager.AppSettings["SVMSGUILink"];
                        var auditDateStr = targetDate.ToString("MM/dd/yyyy");
                        // Generate HTML rows for each company
                        StringBuilder rowsBuilder = new StringBuilder();
                        //foreach (DataRow item in dt.Rows)
                        //{
                        //    rowsBuilder.Append("<tr>");
                        //    rowsBuilder.Append($"<td>{item["CompanyName"]}</td>");
                        //    //auditDateStr = item["AuditDate"] != DBNull.Value
                        //    //                 ? Convert.ToDateTime(item["AuditDate"]).ToString("MM/dd/yyyy")
                        //    //                 : "N/A";
                        //    //rowsBuilder.Append($"<td>{auditDateStr}</td>");

                        //    rowsBuilder.Append("</tr>");
                        //}

                        StringBuilder companyListBuilder = new StringBuilder();
                       // StringBuilder locationBuilder = new StringBuilder();

                        foreach (DataRow item in dt.Rows)
                        {
                            //companyBuilder.Append($"<li>{item["CompanyName"]}</li>");
                            //locationBuilder.Append($"<li>{item["LocationName"]}</li>");
                            string company = item["CompanyName"] != DBNull.Value ? item["CompanyName"].ToString() : "N/A";
                            string location = item["LocationName"] != DBNull.Value ? item["LocationName"].ToString() : "N/A";
                            string missingCount = item["MissingCount"] != DBNull.Value ? item["MissingCount"].ToString() : "0";

                            companyListBuilder.Append("<tr>");
                            companyListBuilder.Append($"<td>{company}</td>");
                            companyListBuilder.Append($"<td>{location}</td>");
                            companyListBuilder.Append($"<td style='text-align:center;'>{missingCount}</td>");
                            companyListBuilder.Append("</tr>");
                        }
                        // Insert company rows into table
                        htmlBody = htmlBody.Replace("#CompanyList", companyListBuilder.ToString());
                        htmlBody = htmlBody.Replace("#Auditdate", auditDateStr);
                        htmlBody = htmlBody.Replace("hrefCode", callbackUrl);
                       // htmlBody = htmlBody.Replace("hrefCode", callbackUrl);


                        // Send email once
                        DAL dal_email = new DAL();
                        dal_email.SendEmailUsingService("ProhibitedMissingAudit", toEmail, CC, BCC, subject, htmlBody, "");
                        LogService.WriteErrorLog("Email sent to: " + toEmail);

                        // Update mail sent status for all rows
                        foreach (DataRow item in dt.Rows)
                        {
                            int companyId = Convert.ToInt32(item["CompanyId"]);
                            int locationId = Convert.ToInt32(item["LocationId"]);
                            if (item["AuditDate"] != DBNull.Value)
                            {
                                DateTime auditDate = Convert.ToDateTime(item["AuditDate"]);
                                dal.UpdateMisingAuditMailSentStatus(companyId, locationId, auditDate);
                            }
                            else
                            {
                                LogService.WriteErrorLog($"Missing AuditDate for CompanyId {companyId}, skipping status update.");
                            }

                           
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(DateTime.Now + " : Error in SendMissingAuditNoticemail(). Exception - " + ex.Message);
            }
        }

    }
}
