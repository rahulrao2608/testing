using EmailDownloadService.Logs;
using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace EmailDownloadService
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        ///         
        static void Main()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[] 
            { 
                new EmailDownloadService() 
            };
            ServiceBase.Run(ServicesToRun);

            EmailDownloadService LED = new EmailDownloadService();
            LED.Process();

            //FetchAllMessages();
        }

        //public static void FetchAllMessages()
        //{
        //    // The client disconnects from the server when being disposed
        //    using (OpenPop.Pop3.Pop3Client client = new OpenPop.Pop3.Pop3Client())
        //    {

        //        string receivingEmailID = "lisainfo@zoho.com";
        //        string hostname = "pop.zoho.com";
        //        int port = 995;
        //        bool useSsl = true;
        //        string username = "lisainfo@zoho.com";
        //        string password = "Member@123";


        //        var wslog = new Log();
        //        //wslog.WriteLog("INF", "Connecting to mail server (" + hostname + ") at port no " + port + sslText);
        //        // Connect to the server
        //        client.Connect(hostname, port, useSsl);
        //        wslog.WriteLog("INF", "Authenticating " + receivingEmailID + " mailbox through POP3");

        //        //OpenPop.Pop3.AuthenticationMethod. AuthMeth = new OpenPop.Pop3.AuthenticationMethod();
        //        //AuthMeth.

        //        // Authenticate ourselves towards the server
        //        client.Authenticate(username, password, OpenPop.Pop3.AuthenticationMethod.UsernameAndPassword);

        //        wslog.WriteLog("INF", receivingEmailID + " successfully authenticated");

        //        // Get the number of messages in the inbox
        //        int messageCount = client.GetMessageCount();

        //        wslog.WriteLog("INF", "Total No of Messages in Inbox(" + receivingEmailID + ")  is " + messageCount);

        //        // We want to download all messages

        //    }
        //}
    }
}
