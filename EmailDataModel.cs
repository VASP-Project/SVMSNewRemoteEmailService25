using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Email_Send_WinService
{
    public class EmailDataModel
    {
        public int Id { get; set; }
        public string ApplicationName { get; set; }
        public string ToMail { get; set; }
        public string CCMail { get; set; }
        public string BCCMail { get; set; }
        public string Subject { get; set; }
        public string MailBody { get; set; }
        public string AttachmentPath { get; set; }
    }

    public class ReminderClearanceData
    {
        public int Id { get; set; }
        public string AuthSignerFirstName { get; set; }
        public string AuthSignerLastName { get; set; }
        public string UserFirstName { get; set; }
        public string UserLastName { get; set; }
        public DateTime? ClearanceDate { get; set; }
        public DateTime? NotificationDate { get; set; }
        public string Email { get; set; }
        public int CompanyId { get; set; }
    }

    public class ReminderAuditData
    {
        public int Id { get; set; }
        public string AuthSignerFirstName { get; set; }
        public string AuthSignerLastName { get; set; }
        public string AuditName { get; set; }
        //public string UserLastName { get; set; }
       // public DateTime ClearanceDate { get; set; }
        public DateTime AuditToDate { get; set; }
        public DateTime AuditFromDate { get; set; }
        public string Email { get; set; }
        public int CompanyId { get; set; }
    }

    public class ReminderNovData
    {
        public int CitationId { get; set; }
        public string AuthSignerFirstName { get; set; }
        public string AuthSignerLastName { get; set; }
        public string Email { get; set; }
        public int CompanyId { get; set; }
        public string CompanyName { get; set; }
        public string ViolatorLastName { get; set; }
        public string ViolatorFirstName { get; set; }
        public DateTime RemedialTrainingAssignedDate { get; set; }
        public int NovNo { get; set; }

    }
}
