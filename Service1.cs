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
            // OverDue Reminder Mail (Re-Reminder of overdue)
            SendOverDueRemindeNovMail();
            //
            TrySendProhibitedAuditNotifications();
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
                                string DecryptedFromDisplayName = EncryptDecryptPassword.DecryptText(config.FromDisplayName, Dkey);
                                string DecryptedEmail = EncryptDecryptPassword.DecryptText(config.Username, Dkey);
                                string DecryptedPassword = EncryptDecryptPassword.DecryptText(config.Password, Dkey);
                                string DecryptedHost = EncryptDecryptPassword.DecryptText(config.SMTPHost, Dkey);
                                string DecryptedPort = EncryptDecryptPassword.DecryptText(config.Port, Dkey);
                                string DecryptedSSL = EncryptDecryptPassword.DecryptText(config.SMTPSSL, Dkey);

                                mail.From = new MailAddress(DecryptedFromEmail, DecryptedFromDisplayName);
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
                                    smtp.EnableSsl = bool.Parse(DecryptedSSL);                                    
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
                LogService.WriteErrorLog("SendReminderMail");
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
                        LogService.WriteErrorLog("Email for companies." + compIds);
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


        public void SendOverDueRemindeNovMail()
        {
            try
            {
                LogService.WriteErrorLog("SendOverdueReminderMail");
                DAL_SVMS dal = new DAL_SVMS();
                DataTable dt = dal.GetOverDueReminderNovData("RD");
                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        
                        string CC = "";
                        string BCC = "";


                        List<ReminderNovData> listName = dt.AsEnumerable().Select(m => new ReminderNovData()
                        {
                            CitationId = m.Field<int>("CitationId"),
                            AuthSignerFirstName = m.Field<string>("AuthSignerFirstName"),
                            AuthSignerLastName = m.Field<string>("AuthSignerLastName"),
                           
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
                                    dal_email.SendEmailUsingService("ReminderOverdueCitation", authSignerEmail, CC, BCC, subject, htmlBody, "");

                                }
                                //  SendMail_SVMSReminder(dt.Rows[i]["Email"].ToString(), "", "", subject, htmlBody, Convert.ToInt32(dt.Rows[i]["Id"]));
                                LogService.WriteErrorLog("Email send to authsigners " + authSignerEmail);

                                dal = new DAL_SVMS();
                                dal.UpdateOverDueNovMailSentStatus(Convert.ToInt32(item.CitationId));
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
                    //LogService.WriteErrorLog("ProhibitedMissingAudit email already sent today. Skipping.");
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
                //DateTime now = DateTime.Now;
                //if (now.Hour < 2 || (now.Hour == 2 && now.Minute < 5))
                //{
                //    LogService.WriteErrorLog($"Skipped sending missing audit mail at {now:yyyy-MM-dd HH:mm:ss}. Waiting until after 2:5 AM.");
                //    return;
                //}
                DateTime now = DateTime.Now;

                // Only allow sending between 02:05 AM and 11:59 PM.
                bool isBeforeGracePeriod = now.Hour < 2 || (now.Hour == 2 && now.Minute < 5);
                LogService.WriteErrorLog("{ now: yyyy - MM - dd HH: mm: ss}");
                if (isBeforeGracePeriod)
                {
                    LogService.WriteErrorLog($"Skipped sending missing audit mail at {now:yyyy-MM-dd HH:mm:ss}. Waiting until after 2:05 AM.");
                    return; // ❗ Correct: return means DO NOT SEND
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

        //public void TrySendProhibitedAuditNotifications()
        //{
        //    try
        //    {
        //        DAL_SVMS dal = new DAL_SVMS();

        //        DataTable dt = dal.GetPendingAuditNotifications();

        //        if (dt == null || dt.Rows.Count == 0)
        //        {
        //            LogService.WriteErrorLog(
        //                "No pending prohibited audit notifications found.");

        //            return;
        //        }

        //        LogService.WriteErrorLog(
        //            $"Pending prohibited audit notifications found: {dt.Rows.Count}");

        //        foreach (DataRow row in dt.Rows)
        //        {
        //            try
        //            {
        //                SendProhibitedAuditNotificationEmail(row, dal);
        //            }
        //            catch (Exception ex)
        //            {
        //                LogService.WriteErrorLog(
        //                    "Error processing prohibited audit notification: "
        //                    + ex.Message);
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        LogService.WriteErrorLog(
        //            "Error in TrySendProhibitedAuditNotifications(): "
        //            + ex.Message);
        //    }
        //}


        //public void SendProhibitedAuditNotificationEmail(
        //    DataRow row,
        //    DAL_SVMS dal)
        //{
        //    try
        //    {
        //        string email =
        //            row["Email"] != DBNull.Value
        //                ? row["Email"].ToString()
        //                : "";

        //        if (string.IsNullOrWhiteSpace(email))
        //        {
        //            LogService.WriteErrorLog(
        //                "AuthSigner email is empty for prohibited audit notification.");

        //            return;
        //        }


        //        //--------------------------------------------------------
        //        // Company
        //        //--------------------------------------------------------
        //        string company =
        //            row["CompanyName"] != DBNull.Value
        //                ? row["CompanyName"].ToString()
        //                : "N/A";


        //        //--------------------------------------------------------
        //        // Location
        //        //--------------------------------------------------------
        //        string location =
        //            row["LocationName"] != DBNull.Value
        //                ? row["LocationName"].ToString()
        //                : "N/A";


        //        //--------------------------------------------------------
        //        // Audit number
        //        //--------------------------------------------------------
        //        int auditNo =
        //            row["AuditNo"] != DBNull.Value
        //                ? Convert.ToInt32(row["AuditNo"])
        //                : 0;


        //        //--------------------------------------------------------
        //        // IsStarted
        //        //
        //        // 0 = Audit has not been started
        //        // 1 = Audit has been started but not submitted
        //        //--------------------------------------------------------
        //        bool isStarted =
        //            row["IsStarted"] != DBNull.Value &&
        //            Convert.ToBoolean(row["IsStarted"]);


        //        //--------------------------------------------------------
        //        // Audit cycle date
        //        //--------------------------------------------------------
        //        DateTime notificationDate =
        //            row["NotificationDate"] != DBNull.Value
        //                ? Convert.ToDateTime(row["NotificationDate"])
        //                : DateTime.Now;


        //        //--------------------------------------------------------
        //        // Assembly path
        //        //--------------------------------------------------------
        //        string assemblyPath =
        //            Path.GetDirectoryName(
        //                System.Reflection.Assembly
        //                    .GetExecutingAssembly()
        //                    .Location);


        //        //--------------------------------------------------------
        //        // Select email template
        //        //--------------------------------------------------------
        //        string templateFileName;

        //        if (!isStarted)
        //        {
        //            templateFileName =
        //                "ProhibitedAuditNotStarted.html";
        //        }
        //        else
        //        {
        //            templateFileName =
        //                "ProhibitedAuditOverdue.html";
        //        }


        //        //--------------------------------------------------------
        //        // Template path
        //        //--------------------------------------------------------
        //        string templatePath =
        //            Path.Combine(
        //                assemblyPath,
        //                "EmailTemplate",
        //                templateFileName);


        //        //--------------------------------------------------------
        //        // Read template
        //        //--------------------------------------------------------
        //        string htmlBody;

        //        using (StreamReader sr = new StreamReader(templatePath))
        //        {
        //            htmlBody = sr.ReadToEnd();
        //        }


        //        //--------------------------------------------------------
        //        // Callback URL
        //        //--------------------------------------------------------
        //        string callbackUrl =
        //            ConfigurationManager.AppSettings["SVMSGUILink"];


        //        //--------------------------------------------------------
        //        // Replace common values
        //        //--------------------------------------------------------
        //        htmlBody = htmlBody.Replace(
        //            "#CompanyName",
        //            company);

        //        htmlBody = htmlBody.Replace(
        //            "#LocationName",
        //            location);

        //        htmlBody = htmlBody.Replace(
        //            "#AuditDate",
        //            notificationDate.ToString("MM/dd/yyyy"));

        //        htmlBody = htmlBody.Replace(
        //            "hrefCode",
        //            callbackUrl);


        //        //--------------------------------------------------------
        //        // Audit number
        //        //--------------------------------------------------------
        //        htmlBody = htmlBody.Replace(
        //            "#AuditNo",
        //            auditNo.ToString());


        //        //--------------------------------------------------------
        //        // Subject and email service template name
        //        //--------------------------------------------------------
        //        string subject;
        //        string emailTemplateName;

        //        if (!isStarted)
        //        {
        //            subject =
        //                "Prohibited Audit - Start Your Audit for Today";

        //            emailTemplateName =
        //                "ProhibitedAuditNotStarted";
        //        }
        //        else
        //        {
        //            subject =
        //                $"Prohibited Audit - Audit {auditNo} Overdue";

        //            emailTemplateName =
        //                "ProhibitedAuditOverdue";
        //        }


        //        //--------------------------------------------------------
        //        // Send email
        //        //--------------------------------------------------------
        //        string CC = "";
        //        string BCC = "";

        //        DAL dalEmail = new DAL();

        //        dalEmail.SendEmailUsingService(
        //            emailTemplateName,
        //            email,
        //            CC,
        //            BCC,
        //            subject,
        //            htmlBody,
        //            "");


        //        //--------------------------------------------------------
        //        // Email sent successfully
        //        // Insert notification log only after successful send
        //        //--------------------------------------------------------
        //        LogService.WriteErrorLog(
        //            $"Prohibited Audit notification email sent. " +
        //            $"Company={company}, " +
        //            $"Location={location}, " +
        //            $"AuditNo={auditNo}, " +
        //            $"IsStarted={isStarted}, " +
        //            $"Email={email}");


        //        dal.InsertProhibitedAuditNotificationLog(
        //            Convert.ToInt32(row["CompanyId"]),
        //            company,
        //            Convert.ToInt32(row["LocationId"]),
        //            location,
        //            auditNo,
        //            "MissedAudit",
        //            notificationDate,
        //            email);
        //    }
        //    catch (Exception ex)
        //    {
        //        LogService.WriteErrorLog(
        //            "Error in SendProhibitedAuditNotificationEmail(): "
        //            + ex.Message);
        //    }
        //}


        public void TrySendProhibitedAuditNotifications()
        {
            try
            {
                DAL_SVMS dal = new DAL_SVMS();

                DataTable dt = dal.GetPendingAuditNotifications();

                if (dt == null || dt.Rows.Count == 0)
                {
                    LogService.WriteErrorLog(
                        "No pending prohibited audit notifications found.");

                    return;
                }

                LogService.WriteErrorLog(
                    $"Pending prohibited audit notifications found: {dt.Rows.Count}");

                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        SendProhibitedAuditNotificationEmail(row, dal);
                    }
                    catch (Exception ex)
                    {
                        LogService.WriteErrorLog(
                            "Error processing prohibited audit notification: "
                            + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(
                    "Error in TrySendProhibitedAuditNotifications(): "
                    + ex.Message);
            }
        }


        public void SendProhibitedAuditNotificationEmail(
            DataRow row,
            DAL_SVMS dal)
        {
            try
            {
                //--------------------------------------------------------
                // Email
                //--------------------------------------------------------
                string email =
                    row["Email"] != DBNull.Value
                        ? row["Email"].ToString()
                        : "";

                if (string.IsNullOrWhiteSpace(email))
                {
                    LogService.WriteErrorLog(
                        "AuthSigner email is empty for prohibited audit notification.");

                    return;
                }


                //--------------------------------------------------------
                // Company
                //--------------------------------------------------------
                string company =
                    row["CompanyName"] != DBNull.Value
                        ? row["CompanyName"].ToString()
                        : "N/A";


                //--------------------------------------------------------
                // Location
                //--------------------------------------------------------
                string location =
                    row["LocationName"] != DBNull.Value
                        ? row["LocationName"].ToString()
                        : "N/A";


                //--------------------------------------------------------
                // Audit number
                //--------------------------------------------------------
                int auditNo =
                    row["AuditNo"] != DBNull.Value
                        ? Convert.ToInt32(row["AuditNo"])
                        : 0;

                //--------------------------------------------------------
                //Overdue Time 
                //--------------------------------------------------------
                TimeSpan overdueTime =
                            row["OverdueTime"] != DBNull.Value
                                ? (TimeSpan)row["OverdueTime"]
                                : TimeSpan.Zero;

                //--------------------------------------------------------
                // Notification Type
                //
                // Reminder = upcoming audit reminder
                // Overdue  = audit overdue
                //--------------------------------------------------------
                string notificationType =
                    row["NotificationType"] != DBNull.Value
                        ? row["NotificationType"].ToString()
                        : "";


                //--------------------------------------------------------
                // Notification date
                //--------------------------------------------------------
                DateTime notificationDate =
                    row["NotificationDate"] != DBNull.Value
                        ? Convert.ToDateTime(row["NotificationDate"])
                        : DateTime.Now;

                //--------------------------------------------------------
                // Audit date
                //
                // Business date of the audit cycle.
                // Audit 3 is checked after midnight but belongs
                // to the previous audit cycle date.
                //--------------------------------------------------------
                DateTime auditDate =
                    row["AuditDate"] != DBNull.Value
                        ? Convert.ToDateTime(row["AuditDate"])
                        : notificationDate;


                //--------------------------------------------------------
                // Assembly path
                //--------------------------------------------------------
                string assemblyPath =
                    Path.GetDirectoryName(
                        System.Reflection.Assembly
                            .GetExecutingAssembly()
                            .Location);


                //--------------------------------------------------------
                // Template / Subject
                //--------------------------------------------------------
                string templateFileName;
                string emailTemplateName;
                string subject;

                

                //--------------------------------------------------------
                // REMINDER
                //--------------------------------------------------------
                if (notificationType.Equals(
                        "Reminder",
                        StringComparison.OrdinalIgnoreCase))
                {
                    templateFileName =
                        "ProhibitedAuditNotStarted.html";

                    emailTemplateName =
                        "ProhibitedAuditNotStarted";

                    //subject =
                       // $"Prohibited Audit - Upcoming Audit {auditNo}";
                }


                //--------------------------------------------------------
                // OVERDUE
                //--------------------------------------------------------
                else if (notificationType.Equals(
                             "Overdue",
                             StringComparison.OrdinalIgnoreCase))
                {
                    templateFileName =
                        "ProhibitedAuditOverdue.html";

                    emailTemplateName =
                        "ProhibitedAuditOverdue";

                   // subject =
                       // $"Prohibited Audit - Audit {auditNo} Overdue";
                }


                //--------------------------------------------------------
                // UNKNOWN TYPE
                //--------------------------------------------------------
                else
                {
                    LogService.WriteErrorLog(
                        $"Unknown prohibited audit notification type: " +
                        $"{notificationType}. " +
                        $"Company={company}, " +
                        $"Location={location}, " +
                        $"AuditNo={auditNo}");

                    return;
                }


                DAL dalEmail = new DAL();

                List<EmailConfigModel> emailConfigs =
                    dalEmail.GetEmailConfigs(string.Empty);

                EmailConfigModel emailConfig =
                    emailConfigs.FirstOrDefault(
                        x => x.ApplicationName.Equals(
                            emailTemplateName,
                            StringComparison.OrdinalIgnoreCase));

                
                
                if (emailConfig == null || string.IsNullOrWhiteSpace(emailConfig.Subject))
                {
                    LogService.WriteErrorLog(
                        $"Email subject configuration not found for {emailTemplateName}");
                    return;
                }

                // Subject and Overdue time replace
                string formattedOverdueTime =
                    DateTime.Today
                        .Add(overdueTime)
                        .ToString(@"HH\:mm");

                                subject = emailConfig.Subject
                                    .Replace("#", auditNo.ToString())
                                    .Replace("Time", formattedOverdueTime);

                //--------------------------------------------------------
                // Template path
                //--------------------------------------------------------
                string templatePath =
                    Path.Combine(
                        assemblyPath,
                        "EmailTemplate",
                        templateFileName);


                //--------------------------------------------------------
                // Check template exists
                //--------------------------------------------------------
                if (!File.Exists(templatePath))
                {
                    LogService.WriteErrorLog(
                        $"Prohibited audit email template not found: " +
                        $"{templatePath}");

                    return;
                }


                //--------------------------------------------------------
                // Read template
                //--------------------------------------------------------
                string htmlBody;

                using (StreamReader sr =
                       new StreamReader(templatePath))
                {
                    htmlBody = sr.ReadToEnd();
                }


                //--------------------------------------------------------
                // Callback URL
                //--------------------------------------------------------
                string callbackUrl =
                    ConfigurationManager.AppSettings["SVMSGUILink"];


                //--------------------------------------------------------
                // Replace template values
                //--------------------------------------------------------
                htmlBody = htmlBody.Replace(
                    "#CompanyName",
                    company);

                htmlBody = htmlBody.Replace(
                    "#LocationName",
                    location);

                htmlBody = htmlBody.Replace(
                    "#AuditDate",
                    auditDate.ToString("MM/dd/yyyy"));

                htmlBody = htmlBody.Replace(
                    "#AuditNo",
                    auditNo.ToString());

                htmlBody = htmlBody.Replace(
                    "hrefCode",
                    callbackUrl);


                //--------------------------------------------------------
                // Send email
                //--------------------------------------------------------
                string CC = "";
                string BCC = "";

                //DAL dalEmail = new DAL();

                
                dalEmail.SendEmailUsingService(
                    emailTemplateName,
                    email,
                    CC,
                    BCC,
                    subject,
                    htmlBody,
                    "");


                //--------------------------------------------------------
                // Email successfully sent
                //--------------------------------------------------------
                LogService.WriteErrorLog(
                    $"Prohibited Audit notification email sent. " +
                    $"Type={notificationType}, " +
                    $"Company={company}, " +
                    $"Location={location}, " +
                    $"AuditNo={auditNo}, " +
                    $"Email={email}");


                //--------------------------------------------------------
                // Insert notification log
                //
                // IMPORTANT:
                // Store exactly the same type returned by SP.
                //
                // Reminder -> Reminder
                // Overdue  -> Overdue
                //--------------------------------------------------------
                dal.InsertProhibitedAuditNotificationLog(
                    Convert.ToInt32(row["CompanyId"]),
                    company,
                    Convert.ToInt32(row["LocationId"]),
                    location,
                    auditNo,
                    notificationType,
                    notificationDate,
                    email);
            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(
                    "Error in SendProhibitedAuditNotificationEmail(): "
                    + ex.Message);
            }
        }
    
    
        
        public void TrySendWeeklyAwsBurdenNotification()
        {
            try
            {
                DAL_SVMS dal = new DAL_SVMS();

                DataTable dt =
                    dal.GetPendingWeeklyAwsBurdenNotifications("AWS");

                if (dt == null || dt.Rows.Count == 0)
                {
                    LogService.WriteErrorLog(
                        "No pending weekly AWS burden notifications found.");
                    return;
                }

                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        SendWeeklyAwsBurdenNotification(row, dal);
                    }
                    catch (Exception ex)
                    {
                        LogService.WriteErrorLog(
                            "Error processing weekly AWS notification for " +
                            "SchedulerInputId " +
                            row["SchedulerInputId"] + ": " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                LogService.WriteErrorLog(
                    DateTime.Now +
                    " : Error in TrySendWeeklyAwsBurdenNotification(): " +
                    ex.Message);
            }
        }


        private void SendWeeklyAwsBurdenNotification(
            DataRow item,
            DAL_SVMS dal)
        {
            int schedulerInputId =
                Convert.ToInt32(item["SchedulerInputId"]);

            string toEmail =
                item["RecipientEmails"] != DBNull.Value
                    ? item["RecipientEmails"].ToString()
                    : string.Empty;

            if (string.IsNullOrWhiteSpace(toEmail))
            {
                LogService.WriteErrorLog(
                    "No AWS notification recipients found for SchedulerInputId " +
                    schedulerInputId);
                return;
            }

            string subject = "AWS Weekly Inspection Burden Not Met";

            DateTime scheduleStartDate =
                Convert.ToDateTime(item["ScheduleStartDate"]);

            DateTime scheduleLastDate =
                Convert.ToDateTime(item["ScheduleLastDate"]);

            string weeklyBurden =
                Convert.ToDecimal(item["WeeklyBurden"])
                    .ToString("0.##");

            string totalDuration =
                item["TotalDuration"] != DBNull.Value
                    ? item["TotalDuration"].ToString()
                    : "0:00";

            string callbackUrl =
                ConfigurationManager.AppSettings["SVMSGUILink"];

            string htmlBody = string.Empty;

            string assemblyPath =
                Path.GetDirectoryName(
                    System.Reflection.Assembly.GetExecutingAssembly().Location);

            string templatePath = Path.Combine(
                assemblyPath,
                "EmailTemplate",
                "AwsWeeklyBurden.html");

            using (StreamReader sr = new StreamReader(templatePath))
            {
                htmlBody = sr.ReadToEnd();
            }

            htmlBody = htmlBody.Replace(
                "#ScheduleStartDate",
                scheduleStartDate.ToString("MM/dd/yyyy"));

            htmlBody = htmlBody.Replace(
                "#ScheduleLastDate",
                scheduleLastDate.ToString("MM/dd/yyyy"));

            htmlBody = htmlBody.Replace(
                "#WeeklyBurden",
                weeklyBurden);

            htmlBody = htmlBody.Replace(
                "#TotalDuration",
                totalDuration);

            htmlBody = htmlBody.Replace(
                "hrefCode",
                callbackUrl ?? string.Empty);

            string CC = string.Empty;
            string BCC = string.Empty;

            DAL dalEmail = new DAL();

            // This call must report failure by throwing an exception
            // if the email could not be sent.
            dalEmail.SendEmailUsingService(
                "AwsWeeklyBurden",
                toEmail,
                CC,
                BCC,
                subject,
                htmlBody,
                "");

            LogService.WriteErrorLog(
                "Weekly AWS burden email processed successfully for " +
                "SchedulerInputId " + schedulerInputId +
                ". Recipients: " + toEmail);

            // Update the sent flag only after the email call succeeds.
            DataTable updateResult =
                dal.GetPendingWeeklyAwsBurdenNotifications(
                    "UC",
                    schedulerInputId);

            if (updateResult != null &&
                updateResult.Rows.Count > 0 &&
                Convert.ToInt32(updateResult.Rows[0]["RowsUpdated"]) == 1)
            {
                LogService.WriteErrorLog(
                    "Weekly AWS notification status updated for SchedulerInputId " +
                    schedulerInputId);
            }
            else
            {
                LogService.WriteErrorLog(
                    "Email processing completed, but weekly AWS notification " +
                    "status was not updated for SchedulerInputId " +
                    schedulerInputId);
            }
        }
    }
}
