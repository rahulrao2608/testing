using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Configuration;
namespace EmailDownloadService.Logs
{
    public class Log
    {
        public void WriteLog(string msgtype, string description)
        {
            string logPath = ConfigurationManager.AppSettings["LogPath"].ToString();
            string infLog = ConfigurationManager.AppSettings["INFLog"].ToString();
            string errLog = ConfigurationManager.AppSettings["ERRLog"].ToString();
            if ((msgtype == "INF" && infLog == "True") || (msgtype == "ERR" && errLog == "True"))
            {
                if (!Directory.Exists(logPath))
                {
                    Directory.CreateDirectory(logPath);
                }
                string filename = "EPSONEmail_Log_" + DateTime.Now.ToString("dd-MM-yyyy") + ".txt";
                string filepath = logPath + filename;
                if (File.Exists(filepath))
                {
                    using (StreamWriter writer = new StreamWriter(filepath, true))
                    {
                        writer.WriteLine(DateTime.Now + "\t" + msgtype + "\t" + description);
                    }
                }
                else
                {
                    StreamWriter writer = File.CreateText(filepath);
                    writer.WriteLine(DateTime.Now + "\t" + msgtype + "\t" + description);
                    writer.Close();
                }
            }
            
           

        }
    }
}
