using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;


namespace EmailDownloadService.Models
{
    
    public class Email
    {
        
        public int EmailId { get; set; }
        
        public long LisaID { get; set; }       
        
        public int CustomerID { get; set; }        

        public string MailFrom { get; set; }

        public string MailTo { get; set; }

        public string MailCC { get; set; }
           
        public string MailBCC { get; set; }
               
        public string MailSubject { get; set; }
        
        public string MailBody { get; set; }

        public DateTime? LastUpdateDate { get; set; }

        public string LastUpdatedBy { get; set; }

        public long AutoResponseID { get; set; }

        public string AutoResponseCode { get; set; }
    }
}
