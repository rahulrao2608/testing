using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using EmailDownloadService.Logs;
using EmailDownloadService.Services;
using System.Timers;
using System.ComponentModel;
using System.Configuration;

namespace EmailDownloadService
{
    public partial class EmailDownloadService : ServiceBase
    {
        private ushort Interval;
        private System.Timers.Timer timer;
        //private BackgroundWorker worker;
        Log wslog = new Log();

        public EmailDownloadService()
        {
            InitializeComponent();
            //worker = new BackgroundWorker();
            //worker.DoWork += worker_DoWork;
        }

        protected override void OnStart(string[] args)
        {
            try
            {
                if (!ushort.TryParse(ConfigurationManager.AppSettings["Interval"].ToString(), out Interval))
                {
                    Interval = 30;
                    wslog.WriteLog("INF", "------Value is not within allowed range (1 to 65536) in the configuration file, service has set its default value to {1} Minute(s).-------");
                }
                timer = new System.Timers.Timer();
                timer.Enabled = true;
                timer.Elapsed += new ElapsedEventHandler(IntervalTimer_Elapsed);
                timer.Interval = ((Convert.ToDouble(Interval) * 60000));
                timer.Start();
                wslog.WriteLog("INF", "------------Email Downloader Service start time set to " + Interval + " minute(s).---------------");
            }
            catch (Exception ex)
            {
                wslog.WriteLog("ERR",  ex.Message.ToString());
            }
        }

        protected override void OnStop()
        {
            //timer.Stop();
            wslog.WriteLog("INF", "------------Attempting to stop service.---------------");
            wslog.WriteLog("INF", "------------Service stopped successfully.---------------");
        }

        private void IntervalTimer_Elapsed(object state, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                timer.Enabled = false;
                Process();
                wslog.WriteLog("INF", "-----------------Email download complete.---------------");
                timer.Enabled = true;
            }
            catch (Exception ex)
            {
                wslog.WriteLog("ERR", ex.ToString().ToString());
            }
            finally
            {
                timer.Enabled = true;
            }

        }

        //void worker_DoWork(object sender, DoWorkEventArgs e)
        //{
        //    // Do the thing that needs doing every few minutes...
        //    // (Omitted for simplicity is sentinel logic to prevent re-entering
        //    //  DoWork() if the previous "tick" has for some reason not completed.)
           
        //}

        public void Process()
        {
            try
            {                
                var emailService = new EmailServices();
                var credentials = emailService.GetCredentials();
                emailService.ReadPop3Emails(credentials);                
                emailService.ExecuteScheduledJob("Processrawdata");
                emailService.ExecuteScheduledJob("LoadAgentQueue", 1);
                emailService.SendAutoResponseMail();
            }
            catch (Exception ex)
            {
                wslog.WriteLog("ERR", ex.Message.ToString());
            }
        }
    }
}
