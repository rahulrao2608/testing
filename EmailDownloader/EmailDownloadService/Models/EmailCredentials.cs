using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EmailDownloadService.Models
{
    public class EmailCredentials
    {
        public int? SmtpID { get; set; }
        public string Title { get; set; }
        public string SmtpFromAddress { get; set; }
        public string SmtpHost { get; set; }
        public int? SmtpPort { get; set; }
        public string SmtpTimeOut { get; set; }
        public bool? SmtpEnableSSL { get; set; }
        public bool? SmtpIsAnonymous { get; set; }
        public string SmtpUser { get; set; }
        public string SmtpPassword { get; set; }
        public bool? Active { get; set; }
        public bool? Default { get; set; }
        public bool? ReceivingMail { get; set; }
        public int? MaxAttachableSizeMB { get; set; }
        public string EmailProtocol { get; set; }
        public string OutHost { get; set; }
        public int OutPort { get; set; }
        public string OutEmailProtocol { get; set; }
    }
}
