using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Email_Send_WinService
{
    public class EmailConfigModel
    {
        public string ApplicationName { get; set; }
        public string FromMail { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public string SMTPHost { get; set; }
        public int SMTPPort { get; set; }
        public string Port { get; set; }
    }
}
