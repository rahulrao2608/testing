using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Net.Mail;
using System.Net.Mime;
using System.Runtime.InteropServices;
using System.Text;
using EmailDownloadService.Logs;
using EmailDownloadService.Models;
using EmailDownloadService.Repository;
using EmailDownloadService.Utilites;
using S22.Pop3;
using System.Web;
using OpenPop.Mime;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using HtmlAgilityPack;
using System.Text.RegularExpressions;


namespace EmailDownloadService.Services
{
    public class EmailServices
    {
        static Dictionary<int,string> allMessagesUID;
        public IEnumerable<EmailCredentials> GetCredentials()
        {
            var wslog = new Log();
            wslog.WriteLog("INF", "Function call for reading email credentials from database");
            var emailRepository = new EmailRepository();
            var repository = emailRepository.GetEmailInformation(InformationTypeEnum.EmailCredentials);
            IEnumerable<EmailCredentials> credentials = null;
            try
            {
                if (repository.ResultTable != null)
                {
                    if (repository.ResultTable.Rows.Count > 0)
                    {
                        credentials = repository.ResultTable.Rows.Cast<DataRow>().Where(x => Convert.ToBoolean(x["Active"]) && Convert.ToBoolean(x["ReceiveMail"])).Select(myRow => new EmailCredentials
                        {
                            SmtpFromAddress = Convert.ToString(myRow["SmtpFromAddress"]),
                            SmtpHost = Convert.ToString(myRow["SmtpHost"]),
                            SmtpPort = Convert.ToInt32(myRow["SmtpPort"]),
                            SmtpUser = Convert.ToString(myRow["SmtpUser"]),
                            SmtpEnableSSL = Convert.ToBoolean(myRow["SmtpEnableSSL"]),
                            SmtpPassword = Convert.ToString(myRow["SmtpPassword"]),
                            EmailProtocol = Convert.ToString(myRow["EmailProtocol"]),
                            ReceivingMail = Convert.ToBoolean(myRow["ReceiveMail"]),
                            MaxAttachableSizeMB = Convert.ToInt32(myRow["MaxAttachmentSizeMB"]),
                            OutHost = Convert.ToString(myRow["OutHost"]),
                            OutPort = Convert.ToInt32(myRow["OutPort"]),
                            OutEmailProtocol = Convert.ToString(myRow["OutEmailProtocol"])
                        });
                    }
                }
                else
                {
                    throw new NoNullAllowedException("No Email credentials found.");
                }
            }
            catch (Exception ex)
            {
                wslog.WriteLog("ERR", ex.Message.ToString());
            }
            return credentials;
        }

        public IEnumerable<EmailCredentials> GetCredentials(string ReceivingEMailID)
        {
            var wslog = new Log();
            wslog.WriteLog("INF", "Function call for reading email credentials from database");
            var emailRepository = new EmailRepository();
            var repository = emailRepository.GetEmailInformation(InformationTypeEnum.EmailCredentials);
            IEnumerable<EmailCredentials> credentials = null;
            try
            {
                if (repository.ResultSet.Tables[0] != null)
                {
                    if (repository.ResultSet.Tables[0].Rows.Count > 0)
                    {
                        credentials = repository.ResultSet.Tables[0].Rows.Cast<DataRow>().Where(x => Convert.ToBoolean(x["Active"]) && Convert.ToBoolean(x["ReceiveMail"])).Select(myRow => new EmailCredentials
                        {
                            SmtpFromAddress = Convert.ToString(myRow["SmtpFromAddress"]),
                            SmtpHost = Convert.ToString(myRow["SmtpHost"]),
                            SmtpPort = Convert.ToInt32(myRow["SmtpPort"]),
                            SmtpUser = Convert.ToString(myRow["SmtpUser"]),
                            SmtpEnableSSL = Convert.ToBoolean(myRow["SmtpEnableSSL"]),
                            SmtpPassword = Convert.ToString(myRow["SmtpPassword"]),
                            EmailProtocol = Convert.ToString(myRow["EmailProtocol"]),
                            ReceivingMail = Convert.ToBoolean(myRow["ReceiveMail"]),
                            MaxAttachableSizeMB = Convert.ToInt32(myRow["MaxAttachmentSizeMB"]),
                            OutHost = Convert.ToString(myRow["OutHost"]),
                            OutPort = Convert.ToInt32(myRow["OutPort"]),
                            OutEmailProtocol = Convert.ToString(myRow["OutEmailProtocol"])

                        });
                    }
                }
                else
                {
                    throw new NoNullAllowedException("No Email credentials found.");
                }
            }
            catch (Exception ex)
            {
                wslog.WriteLog("ERR", ex.Message.ToString());
            }
            return credentials;
        }

        //public void ReadExhangeEmails(IEnumerable<EmailCredentials> emailCredentials)
        //{
        //    var attachmentpath = ConfigurationManager.AppSettings["AttachmentUrl"];
        //    var emailImagesFilePath = ConfigurationManager.AppSettings["ImageFilePath"];
        //    var emailImagesWebUrl = ConfigurationManager.AppSettings["ImageWebUrl"];
        //    var applicationImagesWebUrl = ConfigurationManager.AppSettings["ApplicationImagePath"];
        //    var wslog = new Log();
        //    var emailRepository = new EmailRepository();
        //    try
        //    {
                
        //        foreach (var credentials in emailCredentials.Where(x => x.EmailProtocol == "Exchange"))
        //        {
        //            var GetRemovableText = emailRepository.GetRemovableText(credentials.SmtpFromAddress);
        //            string sslText = "";
        //            bool useSSL = false;
        //            if (credentials.SmtpEnableSSL != null && credentials.SmtpEnableSSL.Value)
        //            {
        //                sslText = "(using SSL)";
        //                useSSL = true;
        //            }


        //            wslog.WriteLog("INF",
        //                "Connecting to mail server (" + credentials.SmtpHost + ") at port no " + credentials.SmtpPort +
        //                sslText);
        //            wslog.WriteLog("INF", "Authenticating " + credentials.SmtpFromAddress + " mailbox through Exchange AutoDiscoverUrl");
        //            //var pop3Client = new Pop3Client(credentials.SmtpHost, Convert.ToInt32(credentials.SmtpPort), credentials.SmtpUser, credentials.SmtpPassword, S22.Pop3.AuthMethod.Login, credentials.SmtpEnableSSL.Value);
                    
        //            //Email exchange starts here

        //            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010)
        //            {
        //                Credentials = new WebCredentials(credentials.SmtpUser, credentials.SmtpPassword)
        //            };

        //            //to add logic for itemview
        //            service.AutodiscoverUrl(credentials.SmtpFromAddress, RedirectionUrlValidationCallback);

        //            wslog.WriteLog("INF", credentials.SmtpFromAddress + " successfully authenticated");
        //            //To get Last email download date
        //            wslog.WriteLog("INF", "Reading" + InformationTypeEnum.LastEmailDownloadDate + "from database");

        //            var getlastDownloadDate =
        //                emailRepository.GetEmailInformation(InformationTypeEnum.LastEmailDownloadDate, credentials.SmtpFromAddress);
        //            if (getlastDownloadDate.OperationStatus == "Error")
        //            {
        //                throw new Exception(
        //                    "Last email download date not found; error at EmailRepository --> GetEmailInformation()");
        //            }
        //            var strlastDownloadDate =
        //                Convert.ToString(getlastDownloadDate.ResultTable.Rows[0]["LastDownloadDate"]);
        //            if (string.IsNullOrEmpty(Convert.ToString(strlastDownloadDate)))
        //            {
        //                wslog.WriteLog("INF",
        //                    "Table is empty, setting last download date to " + DateTime.MinValue.ToString());
        //                strlastDownloadDate = DateTime.MinValue.ToString();
        //            }

        //            var lastDownloadDate = Convert.ToDateTime(strlastDownloadDate);
        //            //
        //            if (credentials.MaxAttachableSizeMB == null || Convert.ToInt32(credentials.MaxAttachableSizeMB) == 0)
        //            {
        //                wslog.WriteLog("INF", "Max attachment size not found, setting default size : 25mb");
        //                credentials.MaxAttachableSizeMB = 25; // Default size for attachment 25mb
        //            }

        //            var maxAttachmentSizeinKbyte = credentials.MaxAttachableSizeMB * 1024;
        //            var euroRawUpdateFlag = "N";
        //            var flag = "";

        //            var getLastMessageId = emailRepository.GetEmailInformation(InformationTypeEnum.MessageId);
        //            if (getLastMessageId.OperationStatus == "Error")
        //            {
        //                throw new Exception(
        //                    "Last MessageID not found; error at EmailRepository --> GetEmailInformation()");
        //            }
        //            var lastMessageId = Convert.ToInt32(getLastMessageId.ResultTable.Rows[0]["MaxMessageId"]);
                    
        //            //var inbox = service.FindItems(WellKnownFolderName.Inbox, new ItemView(100));
        //            SearchFilter filter = new SearchFilter.IsGreaterThan(ItemSchema.DateTimeReceived, lastDownloadDate);
        //            ItemView iview = new ItemView(10);
        //            iview.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Ascending);

        //            var inbox = service.FindItems(WellKnownFolderName.Inbox, filter, iview);

        //            foreach (EmailMessage item in inbox.Items.Where(x => Convert.ToDateTime(x.DateTimeSent) > lastDownloadDate))
        //            {
        //                try
        //                {
        //                    item.Load();
        //                    if (item.DateTimeSent == null)
        //                    {
        //                        continue;
        //                    }

        //                    string threadID = string.Empty;
        //                    wslog.WriteLog("INF", "Mail Date Extracted form Mail Box :" + item.DateTimeSent);
        //                    var currentEmailDate = Convert.ToDateTime(item.DateTimeSent);
        //                    wslog.WriteLog("INF", "Reading last " + InformationTypeEnum.MessageId);
        //                    lastMessageId = lastMessageId == 0 ? 1 : lastMessageId + 1;
        //                    const decimal kilobyte = 1024;
        //                    flag = "TEMP";// TO INSERT INTO TEMPORARY TABLE

        //                    wslog.WriteLog("INF", "Reading new emails from " + credentials.SmtpFromAddress + " mail box");
        //                    euroRawUpdateFlag = "Y";

        //                    //var ccList = string.Join(",", item.CcRecipients.Select(x => x.Address.ToString()));
        //                    //var toList = string.Join(",", item.ToRecipients.Select(x => x.Address.ToString()));

        //                    var ccList = string.Join(";", item.CcRecipients.Select(x => x.Address.ToString()));
        //                    var toList = string.Join(";", item.ToRecipients.Select(x => x.Address.ToString()));

        //                    var strMessageBody = item.Body.Text;

        //                    if (!strMessageBody.Contains("EmailThreadStart:") && !strMessageBody.Contains(":EmailThreadEnd"))
        //                    {
        //                        var emailThreadID = emailRepository.GetThreadID(credentials.SmtpFromAddress, item.From.Address, toList, ccList, item.Subject, "");
        //                        threadID = Convert.ToString(emailThreadID.ResultTable.Rows[0]["ThreadID"]);
        //                        strMessageBody = strMessageBody.Replace("</body>", "<div id=\"LisaMessageThreadIdentifier\" style=\"visibility:hidden\">EmailThreadStart:" + threadID + ":EmailThreadEnd</div></body>");
        //                    }
        //                    else
        //                    {
        //                        int startIndex = strMessageBody.IndexOf("EmailThreadStart:");
        //                        int endIndex = strMessageBody.IndexOf(":EmailThreadEnd");

        //                        if (endIndex > startIndex)
        //                            threadID = strMessageBody.Substring(startIndex, endIndex - startIndex);
        //                        threadID = threadID.Replace("EmailThreadStart:", "");
        //                        //strMessageBody = strMessageBody.Replace("</body>", "<div id=\"MessageThreadIdentifier\">" + threadID + "</div></body>");
        //                    }
        //                    strMessageBody = ExtractRecentMailContent(strMessageBody, "EmailThreadStart:" + threadID + ":EmailThreadEnd");


        //                    foreach (FileAttachment attachment in item.Attachments)
        //                    {
        //                        attachment.Load();
        //                        if (attachment.IsInline)
        //                        {
        //                            applicationImagesWebUrl = ConfigurationManager.AppSettings["ApplicationImagePath"];
        //                            var dirImages = new DirectoryInfo(emailImagesFilePath + credentials.SmtpFromAddress.Split('@')[0] + "_" + currentEmailDate.ToString("yyyyMMddHHmmss"));
        //                            if (!dirImages.Exists)
        //                                wslog.WriteLog("INF", "Creating image directory for " + credentials.SmtpFromAddress + " mail box");
        //                            dirImages.Create();

        //                            wslog.WriteLog("INF", "Saving attachments for " + credentials.SmtpFromAddress + " mail box");
        //                            var filenameByte = dirImages.FullName + DateTime.Now.ToString("yyyyMMddHHmmssffffff") +
        //                                                  attachment.Name;
        //                            // to correct logic for multiple inline images in body                                    
        //                            applicationImagesWebUrl = emailImagesWebUrl + "\\" + filenameByte.Replace(emailImagesFilePath, "");
        //                            File.WriteAllBytes(filenameByte, attachment.Content);
        //                            item.Body = item.Body.Text.Replace("cid:" + attachment.ContentId, applicationImagesWebUrl);


        //                        }
        //                        else
        //                        {
        //                            var sizeofFile = Math.Round((attachment.Size / kilobyte), 2);
        //                            if (sizeofFile > maxAttachmentSizeinKbyte)
        //                            {
        //                                throw new Exception("Attachment size is greater than Maximum attachment size");
        //                            }
        //                            var dr = new DirectoryInfo(attachmentpath + item.From.Address + "_" + currentEmailDate.ToString("yyyyMMddHHmmss") + @"\");
        //                            if (!dr.Exists)
        //                                wslog.WriteLog("INF", "Creating attachment directory for " + credentials.SmtpFromAddress + " mail box");
        //                            dr.Create();

        //                            wslog.WriteLog("INF", "Saving attachments for " + credentials.SmtpFromAddress + " mail box");
        //                            var filenameByte = dr.FullName + DateTime.Now.ToString("yyyyMMddHHmmssffffff") +
        //                                                  attachment.Name;
        //                            File.WriteAllBytes(filenameByte, attachment.Content);

        //                            var attachmentResponse =
        //                              emailRepository.DownloadAttachments(filenameByte, 0,
        //                                  attachment.Name, sizeofFile, lastMessageId);

        //                            if (attachmentResponse.OperationStatus == "Error")
        //                            {
        //                                throw new Exception("Error while downloading attachments to table ; error at EmailRepository --> DownloadAttachments()");
        //                            }
        //                        }
        //                    }
        //                    //Modify By Kuldeep ---Start---
        //                    foreach (DataRow Dr in GetRemovableText.ResultSet.Tables[0].Rows)
        //                    {
        //                        if ((item.From).ToString().EndsWith(Dr["EmailId"].ToString()))
        //                        {
        //                            item.Body = item.Body.Text.Replace(Dr["Content"].ToString(), " ");

        //                        }
        //                    }

        //                    //End
        //                    //Modify By Kuldeep
        //                    DateTime ReciveTime = item.DateTimeReceived.ToUniversalTime();

        //                    var downloadEmailUtility = emailRepository.DownloadEmail(flag, item.From.Address, toList,
        //                    ccList, item.Subject, strMessageBody, item.InternetMessageId, item.InternetMessageId, threadID, ReciveTime,
        //                     item.From.Name, lastMessageId, 0, credentials.SmtpFromAddress, currentEmailDate);

        //                    if (downloadEmailUtility.OperationStatus == "Error")
        //                    {
        //                        throw new Exception("Error while inserting into TempEmailDownloadEuroRawData; error at EmailRepository --> DownloadEmail");
        //                    }

        //                    //--Start Commented By Bidhan--To eliminate extra DBHit
        //                    //var saveEmailDownloadDateUtility = emailRepository.SaveEmailDownloadDate(currentEmailDate, credentials.SmtpFromAddress);

        //                    //if (saveEmailDownloadDateUtility.OperationStatus == "Error")
        //                    //{
        //                    //    throw new Exception("Error while inserting into EmailDownloadDateMaster; error at EmailRepository --> SaveEmailDownloadDate");
        //                    //}

        //                    //--End Commented By Bidhan--To eliminate extra DBHit


        //                    //to write logic for saving muid
        //                    //var pop3MessageId = emailRepository.InsertPop3MessageGuid(item.Key);
        //                    //if (pop3MessageId.OperationStatus == "Error")
        //                    //{
        //                    //    throw new Exception("Error while inserting into POP3MessageMaster; error at EmailRepository --> InsertPop3MessageGuid");
        //                    //}

        //                    //ends

        //                }
        //                catch (Exception e)
        //                {
        //                    wslog.WriteLog("ERR", e.Message.ToString());
        //                }

        //                //break;
        //            }
        //            wslog.WriteLog("INF", "No new mails present in " + credentials.SmtpFromAddress + " mail box");
        //            try
        //            {
        //                if (euroRawUpdateFlag == "Y")
        //                {
        //                    wslog.WriteLog("INF", "Inserting into EuroRawData from TempEuroRawData");
        //                    flag = "EURO";
        //                    var euroDownloadUtility = emailRepository.DownloadEmail(flag, "", "", ""
        //                           , "", "","","","", DateTime.Now, "", 0, 0, "", DateTime.Today);
        //                    if (euroDownloadUtility.OperationStatus == "Error")
        //                    {
        //                        throw new Exception("Error while inserting into EuroRawData; error at EmailRepository --> DownloadEmail");
        //                    }
        //                    wslog.WriteLog("INF", "Process completed");
        //                }
        //            }
        //            catch (Exception e)
        //            {
        //                wslog.WriteLog("ERR", e.Message.ToString());
        //            }
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        wslog.WriteLog("ERR", e.Message.ToString());
        //    }
        //}

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        public void ReadPop3Emails(IEnumerable<EmailCredentials> emailCredentials)
        {
            var attachmentpath = ConfigurationManager.AppSettings["AttachmentUrl"];
            var emailImagesFilePath = ConfigurationManager.AppSettings["ImageFilePath"];
            var emailImagesWebUrl = ConfigurationManager.AppSettings["ImageWebUrl"];
            var applicationImagesWebUrl = ConfigurationManager.AppSettings["ApplicationImagePath"];
            var wslog = new Log();
            var emailRepository = new EmailRepository();
            try
            {
                foreach (var credentials in emailCredentials.Where(x => x.EmailProtocol == "POP3"))
                {
                    try
                    {
                        var GetRemovableText = emailRepository.GetRemovableText(credentials.SmtpFromAddress);
                        string sslText = "";
                        bool useSSL = false;
                        if (credentials.SmtpEnableSSL != null && credentials.SmtpEnableSSL.Value)
                        {
                            sslText = "(using SSL)";
                            useSSL = true;
                        }

                        //To get Last email download date
                        wslog.WriteLog("INF", "Reading" + InformationTypeEnum.LastEmailDownloadDate + "from database");
                        //
                        var getlastDownloadDate = emailRepository.GetEmailInformation(InformationTypeEnum.LastEmailDownloadDate, credentials.SmtpFromAddress);
                        if (getlastDownloadDate.OperationStatus == "Error")
                        {
                            throw new Exception("Last email download date not found; error at EmailRepository --> GetEmailInformation()");
                        }
                        var strlastDownloadDate = Convert.ToString(getlastDownloadDate.ResultTable.Rows[0]["LastDownloadDate"]);
                        if (string.IsNullOrEmpty(Convert.ToString(strlastDownloadDate)))
                        {
                            wslog.WriteLog("INF", "Table is empty, setting last download date to " + DateTime.MinValue.ToString());
                            strlastDownloadDate = DateTime.MinValue.ToString();
                        }

                        var lastDownloadDate = Convert.ToDateTime(strlastDownloadDate);
                        //
                        if (credentials.MaxAttachableSizeMB == null || Convert.ToInt32(credentials.MaxAttachableSizeMB) == 0)
                        {
                            wslog.WriteLog("INF", "Max attachment size not found, setting default size : 25mb");
                            credentials.MaxAttachableSizeMB = 25; // Default size for attachment 25mb
                        }

                        var maxAttachmentSizeinKbyte = credentials.MaxAttachableSizeMB * 1024;
                        var euroRawUpdateFlag = "N";
                        var flag = "";

                        var getLastMessageId = emailRepository.GetEmailInformation(InformationTypeEnum.MessageId);
                        if (getLastMessageId.OperationStatus == "Error")
                        {
                            throw new Exception("Last MessageID not found; error at EmailRepository --> GetEmailInformation()");
                        }
                        var lastMessageId = Convert.ToInt32(getLastMessageId.ResultTable.Rows[0]["MaxMessageId"]);
                        var inbox = FetchAllMessages(credentials.SmtpFromAddress, credentials.SmtpHost, Convert.ToInt32(credentials.SmtpPort), useSSL, credentials.SmtpUser, credentials.SmtpPassword,Convert.ToBoolean(credentials.SmtpIsAnonymous));
                        if (inbox.Count <= 0)
                        {
                            wslog.WriteLog("INF", "Inbox(" + credentials.SmtpFromAddress + ") is empty");
                        }
                        else
                        {
                            //foreach (var item in inbox.Where(x => x.Headers["Date"] != null && Convert.ToDateTime(x.Headers["Date"]) > lastDownloadDate))
                            foreach (var item in inbox.OrderBy(x => Convert.ToDateTime(x.Value.Headers.DateSent)))
                            {
                                try
                                {
                                    MailMessage objMailMessage = item.Value.ToMailMessage();
                                    var strMessageBody = objMailMessage.Body;
                                    var threadID = string.Empty;
                                    var ccList = string.Join(";", objMailMessage.CC.Select(u => u.Address));
                                    var toList = string.Join(";", objMailMessage.To.Select(u => u.Address));

                                    if (!strMessageBody.Contains("EmailThreadStart:") && !strMessageBody.Contains(":EmailThreadEnd"))
                                    {
                                        var emailThreadID = emailRepository.GetThreadID(credentials.SmtpFromAddress, objMailMessage.From.Address, toList, ccList, objMailMessage.Subject, "");
                                        threadID = Convert.ToString(emailThreadID.ResultTable.Rows[0]["ThreadID"]);
                                        strMessageBody = strMessageBody.Replace("</body>", "<div id=\"LisaMessageThreadIdentifier\" style=\"visibility:hidden\">EmailThreadStart:" + threadID + ":EmailThreadEnd</div></body>");
                                    }
                                    else
                                    {
                                        int startIndex = strMessageBody.IndexOf("EmailThreadStart:");
                                        int endIndex = strMessageBody.IndexOf(":EmailThreadEnd");

                                        if (endIndex > startIndex)
                                            threadID = strMessageBody.Substring(startIndex, endIndex - startIndex);
                                        threadID = threadID.Replace("EmailThreadStart:", "");
                                        //strMessageBody = strMessageBody.Replace("</body>", "<div id=\"MessageThreadIdentifier\">" + threadID + "</div></body>");
                                    }
                                    strMessageBody = ExtractRecentMailContent(strMessageBody, "EmailThreadStart:" + threadID + ":EmailThreadEnd");


                                    if (item.Value.Headers.Date == null)
                                    {
                                        continue;
                                    }
                                    //item.Value.
                                    wslog.WriteLog("INF", "Mail Date Extracted form Mail Box :" + item.Value.Headers.DateSent);
                                    var currentEmailDate = Convert.ToDateTime(item.Value.Headers.DateSent);
                                    wslog.WriteLog("INF", "Reading last " + InformationTypeEnum.MessageId);
                                    lastMessageId = lastMessageId == 0 ? 1 : lastMessageId + 1;
                                    const decimal kilobyte = 1024;
                                    flag = "TEMP";// TO INSERT INTO TEMPORARY TABLE

                                    try
                                    {
                                        if (item.Value.FindAllMessagePartsWithMediaType("image/png").Count > 0 || item.Value.FindAllMessagePartsWithMediaType("image/jpeg").Count > 0)
                                        {
                                            var dirImages = new DirectoryInfo(emailImagesFilePath + credentials.SmtpFromAddress.Split('@')[0] + "_" + currentEmailDate.ToString("yyyyMMddHHmmss"));
                                            if (!dirImages.Exists)
                                                wslog.WriteLog("INF", "Creating image directory for " + credentials.SmtpFromAddress + " mail box");
                                            dirImages.Create();
                                        }

                                        List<MessagePart> objMessagePart = item.Value.FindAllMessagePartsWithMediaType("image/png");
                                        applicationImagesWebUrl = ConfigurationManager.AppSettings["ApplicationImagePath"];
                                        foreach (MessagePart objMess in objMessagePart)
                                        {
                                            var strDateStamp = DateTime.Now.ToString("yyyyMMddHHmmssffffff");
                                            var contentID = String.IsNullOrEmpty(objMess.ContentId) ? "" : objMess.ContentId;
                                            var fileName = objMess.FileName;
                                            if (contentID.Length > 30)
                                                contentID = contentID.Substring(0, 10);
                                            if (fileName.Length > 30)
                                                fileName = fileName.Substring(25, fileName.Length - 25);
                                            if (fileName.Length > 0 && !fileName.Contains("(no name)"))
                                            {
                                                var uniqueContentID = "StartLisacid_" + contentID + "_EndLisacid_" + fileName + "_EndLisaImage";
                                                uniqueContentID = uniqueContentID.Replace('.', '_');
                                                var fileNameWithPath = emailImagesFilePath + credentials.SmtpFromAddress.Split('@')[0] + "_" + currentEmailDate.ToString("yyyyMMddHHmmss") + "\\" + strDateStamp + uniqueContentID + fileName;
                                                var fileImageURL = emailImagesWebUrl + credentials.SmtpFromAddress.Split('@')[0] + "_" + currentEmailDate.ToString("yyyyMMddHHmmss") + "\\" + strDateStamp + uniqueContentID + fileName;
                                                // var fileImageURL = fileNameWithPath;
                                                using (var fileStream = new FileStream(fileNameWithPath, FileMode.Create, FileAccess.Write))
                                                {
                                                    objMess.Save(fileStream);
                                                }

                                                strMessageBody = strMessageBody.Replace("cid:" + objMess.ContentId, fileImageURL);
                                            }
                                        }

                                        objMessagePart.Clear();
                                        objMessagePart = item.Value.FindAllMessagePartsWithMediaType("image/jpeg");
                                        foreach (MessagePart objMess in objMessagePart)
                                        {
                                            var strDateStamp = DateTime.Now.ToString("yyyyMMddHHmmssffffff");
                                            var contentID = String.IsNullOrEmpty(objMess.ContentId) ? "" : objMess.ContentId;
                                            var fileName = objMess.FileName;
                                            if (contentID.Length > 30)
                                                contentID = contentID.Substring(0, 10);
                                            if (fileName.Length > 30)
                                                fileName = fileName.Substring(25, fileName.Length - 25);
                                            if (fileName.Length > 0 && !fileName.Contains("(no name)"))
                                            {
                                                var uniqueContentID = "StartLisacid_" + contentID + "_EndLisacid_" + fileName + "_EndLisaImage";
                                                //var uniqueContentID = "StartLisacid_" + objMess.ContentId + "_EndLisacid_" + objMess.FileName + "_EndLisaImage";
                                                uniqueContentID = uniqueContentID.Replace('.', '_');
                                                var fileNameWithPath = emailImagesFilePath + credentials.SmtpFromAddress.Split('@')[0] + "_" + currentEmailDate.ToString("yyyyMMddHHmmss") + "\\" + strDateStamp + uniqueContentID + fileName;
                                                var fileImageURL = emailImagesWebUrl + credentials.SmtpFromAddress.Split('@')[0] + "_" + currentEmailDate.ToString("yyyyMMddHHmmss") + "\\" + strDateStamp + uniqueContentID + fileName;
                                                //var fileImageURL = fileNameWithPath;
                                                using (var fileStream = new FileStream(fileNameWithPath, FileMode.Create, FileAccess.Write))
                                                {
                                                    objMess.Save(fileStream);
                                                }

                                                strMessageBody = strMessageBody.Replace("cid:" + objMess.ContentId, fileImageURL);
                                            }
                                        }

                                        objMessagePart.Clear();
                                        objMessagePart = item.Value.FindAllMessagePartsWithMediaType("image/gif");
                                        foreach (MessagePart objMess in objMessagePart)
                                        {
                                            var strDateStamp = DateTime.Now.ToString("yyyyMMddHHmmssffffff");
                                            var contentID = String.IsNullOrEmpty(objMess.ContentId) ? "" : objMess.ContentId;
                                            var fileName = objMess.FileName;
                                            if (contentID.Length > 30)
                                                contentID = contentID.Substring(0, 10);
                                            if (fileName.Length > 30)
                                                fileName = fileName.Substring(25, fileName.Length - 25);

                                            if (fileName.Length > 0 && !fileName.Contains("(no name)"))
                                            {
                                                var uniqueContentID = "StartLisacid_" + contentID + "_EndLisacid_" + fileName + "_EndLisaImage";
                                                uniqueContentID = uniqueContentID.Replace('.', '_');
                                                var fileNameWithPath = emailImagesFilePath + credentials.SmtpFromAddress.Split('@')[0] + "_" + currentEmailDate.ToString("yyyyMMddHHmmss") + "\\" + strDateStamp + uniqueContentID + fileName;
                                                var fileImageURL = emailImagesWebUrl + credentials.SmtpFromAddress.Split('@')[0] + "_" + currentEmailDate.ToString("yyyyMMddHHmmss") + "\\" + strDateStamp + uniqueContentID + fileName;
                                                //var fileImageURL = fileNameWithPath;
                                                using (var fileStream = new FileStream(fileNameWithPath, FileMode.Create, FileAccess.Write))
                                                {
                                                    objMess.Save(fileStream);
                                                }

                                                strMessageBody = strMessageBody.Replace("cid:" + objMess.ContentId, fileImageURL);
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        wslog.WriteLog("ERR", ex.Message.ToString());
                                    }

                                    try
                                    {
                                        wslog.WriteLog("INF", "Reading new emails from " + credentials.SmtpFromAddress + " mail box");
                                        euroRawUpdateFlag = "Y";
                                        if (item.Value.FindAllAttachments().Count > 0)
                                        {
                                            foreach (var attachment in item.Value.FindAllAttachments()) //objMailMessage.Attachments
                                            {
                                                if (String.IsNullOrEmpty(attachment.ContentId))
                                                {
                                                    wslog.WriteLog("INF", "Attachment found for MessageId: " + lastMessageId + " >> File Name: " + attachment.FileName);
                                                    if (!string.IsNullOrEmpty(Path.GetExtension(attachment.FileName)))
                                                    {
                                                        bool IsAttachmentAllowed = CheckExcludeAttachmentTypes(Path.GetExtension(attachment.FileName));
                                                        if (IsAttachmentAllowed)
                                                        {
                                                            var sizeofFile = Math.Round((attachment.ContentDescription.Length / kilobyte), 2);
                                                            if (sizeofFile > maxAttachmentSizeinKbyte)
                                                            {
                                                                throw new Exception("Attachment size is greater than Maximum attachment size");
                                                            }

                                                            string fileName = string.Empty;
                                                            fileName = attachment.FileName;
                                                            try
                                                            {
                                                                var dr = new DirectoryInfo(attachmentpath + objMailMessage.From.Address + "_" + currentEmailDate.ToString("yyyyMMddHHmmss") + @"\");
                                                                if (!dr.Exists)
                                                                    wslog.WriteLog("INF", "Creating attachment directory for " + credentials.SmtpFromAddress + " mail box");
                                                                dr.Create();

                                                                Regex rgx = new Regex(@"[^0-9a-zA-Z-_. ,]");
                                                                if (!string.IsNullOrEmpty(fileName) && rgx.IsMatch(fileName))
                                                                {
                                                                    fileName = Regex.Replace(attachment.FileName, @"[^0-9a-zA-Z-_. ,]", "");
                                                                    wslog.WriteLog("INF", "Attachment file contains some invalid characters.");
                                                                    wslog.WriteLog("INF", "Attachment file name has been changed from " + attachment.FileName + " to " + fileName);
                                                                }


                                                                // Creating attachments in directory
                                                                using (var fileStream = new FileStream(dr.FullName + DateTime.Now.ToString("yyyyMMddHHmmssffffff") + fileName, FileMode.Create, FileAccess.Write))
                                                                {
                                                                    wslog.WriteLog("INF", "Saving attachments for " + credentials.SmtpFromAddress + " mail box");
                                                                    attachment.Save(fileStream);

                                                                    var attachmentResponse =
                                                                      emailRepository.DownloadAttachments(fileStream.Name, 0,
                                                                          fileName, sizeofFile, lastMessageId);

                                                                    if (attachmentResponse.OperationStatus == "Error")
                                                                    {
                                                                        throw new Exception("Error while downloading attachments to table ; error at EmailRepository --> DownloadAttachments()");
                                                                    }

                                                                }

                                                            }

                                                            catch (Exception ex)
                                                            {
                                                                wslog.WriteLog("ERR", ex.Message.ToString());
                                                            }
                                                        }
                                                        else
                                                        {
                                                            wslog.WriteLog("INF", "Attachment type " + "'" + Path.GetExtension(attachment.FileName).ToUpper() + "'" + " not allowed.");
                                                        }
                                                    }
                                                    
                                                }
                                               
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        wslog.WriteLog("ERR", ex.Message.ToString());
                                    }

                                    try
                                    {
                                        if (GetRemovableText.ResultTable != null)
                                        {
                                            if (GetRemovableText.ResultTable.Rows.Count > 0)
                                            {
                                                foreach (DataRow Dr in GetRemovableText.ResultTable.Rows)
                                                {
                                                    strMessageBody = strMessageBody.Replace(Dr["Content"].ToString(), " ");
                                                }
                                            }
                                        }

                                        string messageUID = allMessagesUID.SingleOrDefault(x => x.Key == item.Key).Value;
                                        //string messageUID = IDUID.Value;
                                        var downloadEmailUtility = emailRepository.DownloadEmail(flag, objMailMessage.From.Address, toList,
                                        ccList, objMailMessage.Subject, strMessageBody, item.Value.Headers.MessageId, messageUID, threadID, currentEmailDate,
                                         objMailMessage.From.DisplayName, lastMessageId, item.Key, credentials.SmtpFromAddress, currentEmailDate);

                                        if (downloadEmailUtility.OperationStatus == "Error")
                                        {
                                            throw new Exception("Error while inserting into TempEmailDownloadEuroRawData or POP3MessageMaster or EmailDownloadDateMaster; error at EmailRepository --> DownloadEmail");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        wslog.WriteLog("ERR", ex.Message.ToString());
                                    }

                                    try
                                    {
                                        emailRepository.AddAutoResponseRecord(objMailMessage.From.Address, credentials.SmtpFromAddress, objMailMessage.Subject, threadID, currentEmailDate, DateTime.Now.ToString("yyyyMMddHHmmssffffff"));
                                    }
                                    catch (Exception ex)
                                    {
                                        wslog.WriteLog("ERR", ex.Message.ToString());
                                    }
                                }
                                catch (Exception e)
                                {
                                    wslog.WriteLog("ERR", e.Message.ToString());
                                }
                                //break;
                            }

                            wslog.WriteLog("INF", "No new mails present in " + credentials.SmtpFromAddress + " mail box");
                            try
                            {
                                if (euroRawUpdateFlag == "Y")
                                {
                                    wslog.WriteLog("INF", "Inserting into EuroRawData from TempEuroRawData");
                                    flag = "EURO";
                                    var euroDownloadUtility = emailRepository.DownloadEmail(flag, "", "", ""
                                           , "", "", "", "", "", DateTime.Now, "", 0, 0, "", DateTime.Today);
                                    if (euroDownloadUtility.OperationStatus == "Error")
                                    {
                                        throw new Exception("Error while inserting into EuroRawData; error at EmailRepository --> DownloadEmail");
                                    }
                                    wslog.WriteLog("INF", "Process completed");
                                }
                            }
                            catch (Exception e)
                            {
                                wslog.WriteLog("ERR", e.Message.ToString());
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        wslog.WriteLog("ERR", ex.Message.ToString());
                    }

                }

            }
            catch (Exception e)
            {
                wslog.WriteLog("ERR", e.Message.ToString());
            }

        }
        public static Dictionary<int, Message> FetchAllMessages(string receivingEmailID, string hostname, int port, bool useSsl, string username, string password,bool allowAnonymous=true)
        {
            // The client disconnects from the server when being disposed
            using (OpenPop.Pop3.Pop3Client client = new OpenPop.Pop3.Pop3Client())
            {
                var emailRepository = new EmailRepository();               

                var sslText = String.Empty;
                if (useSsl)
                {
                    sslText = " (using SSL)";
                }
                //logs
                var wslog = new Log();
                wslog.WriteLog("INF", "Connecting to mail server (" + hostname + ") at port no " + port + sslText);
                // Connect to the server
                client.Connect(hostname, port, useSsl);
                wslog.WriteLog("INF", "Authenticating " + receivingEmailID + " mailbox through POP3");

                // Authenticate ourselves towards the server
                if (allowAnonymous)
                {
                    client.Authenticate(username, password);
                }
                else
                {                    
                    client.Authenticate(username, password, OpenPop.Pop3.AuthenticationMethod.UsernameAndPassword);
                }

                wslog.WriteLog("INF", receivingEmailID + " successfully authenticated");

                // Get the number of messages in the inbox
                int messageCount = client.GetMessageCount();

                wslog.WriteLog("INF", "Total No of Messages in Inbox(" + receivingEmailID + ")  is " + messageCount);

                // We want to download all messages
                var allMessages = new Dictionary<int, Message>(messageCount);
                allMessagesUID = new Dictionary<int, string>(messageCount);

                // Messages are numbered in the interval: [1, message++Count]
                // Ergo: message numbers are 1-based.
                // Most servers give the latest message the highest number

                var lastPop3Uid = emailRepository.GetEmailInformation(InformationTypeEnum.Pop3MessageId, receivingEmailID);
                if (lastPop3Uid.OperationStatus == "Error")
                {
                    throw new Exception("Last MUId not found; error at EmailRepository --> GetEmailInformation()");
                }
                var id = Convert.ToInt32(lastPop3Uid.ResultTable.Rows[0]["lastMuid"]);
                var lastEmailDownloadDate = Convert.ToDateTime(lastPop3Uid.ResultTable.Rows[0]["LastDownloadDate"]);

                if (messageCount > id)
                {
                    wslog.WriteLog("INF", "New Messages in Inbox(" + receivingEmailID + ")  is " + (messageCount - id));
                    for (int i = messageCount; i > id; i--)
                    {
                        //allMessages.Add(client.GetMessage(i));
                        allMessages.Add(i, client.GetMessage(i));
                        allMessagesUID.Add(i, client.GetMessageUid(i));
                    }
                }
                else
                {
                    var messageExist = false;
                    for (; messageCount > 0; messageCount--)
                    {
                        //allMessages.Add(client.GetMessage(i));
                        if (Convert.ToDateTime(client.GetMessageHeaders(messageCount).DateSent) > lastEmailDownloadDate)
                        {
                            allMessages.Add(messageCount, client.GetMessage(messageCount));
                            allMessagesUID.Add(messageCount, client.GetMessageUid(messageCount));
                            messageExist = true;
                        }
                        else
                            break;
                    }

                    if (messageExist)
                    {
                        emailRepository.UpdatePOP3MessageID(receivingEmailID, messageCount);
                        wslog.WriteLog("INF", "Updated POP3MessageMaster for Inbox(" + receivingEmailID + ").");
                    }
                }
                //if (messageCount > id)
                //{
                    wslog.WriteLog("INF", "Desk Name: " + receivingEmailID + " >> Total Emails On Server: " + messageCount + " | Last MUID: " + id + " | Diff of Total Emails and MUID: " + (messageCount - id));
                    var InsertTotalEmailCountOnServer = emailRepository.InsertTotalEmailCountOnServer(messageCount, Convert.ToInt64(messageCount - id), receivingEmailID);
                    wslog.WriteLog("INF", "Total Email Count details inserted into >> TotalEmailOnServerDetails table.");
                    if (InsertTotalEmailCountOnServer.OperationStatus == "Error")
                    {
                        throw new Exception("Error while Insering Total Email Count available on Server; error at EmailRepository --> InsertTotalEmailCountOnServer()");
                    }
                //}

                // Now return the fetched messages
                return allMessages;
            }
        }
        
        public void SendAutoResponseMail()
        {
            var wslog = new Log();            
            EmailRepository objEmailRepository = new EmailRepository();
            try
            {
                EmailUtilities objEmailUtilities = objEmailRepository.GetAllAutoResponseRecord();
                if (objEmailUtilities != null)
                {
                    foreach (DataRow dr in objEmailUtilities.ResultTable.Rows)
                    {
                        try
                        {
                            EmailPusher objEmailPusher = new EmailPusher();
                            Email objEmail = new Email();
                            objEmail.MailFrom = Convert.ToString(dr["ReceivingMailID"]);
                            objEmail.MailTo = Convert.ToString(dr["FromMailID"]);
                            objEmail.AutoResponseCode = Convert.ToString(dr["AutoResponseID"]);
                            objEmail.MailSubject = "Reference No(" + Convert.ToString(dr["AutoResponseID"]) + ") : " + Convert.ToString(dr["Subject"]);
                            objEmail.MailBody = "<html><body>" + ConfigurationManager.AppSettings["ResponseMessage1"] + "<br/><br/>" + ConfigurationManager.AppSettings["ResponseMessage2"] + "<br/>" + ConfigurationManager.AppSettings["ResponseMessage3"] + " <strong>" + Convert.ToString(dr["AutoResponseID"]) + "</strong>. " + ConfigurationManager.AppSettings["ResponseMessage4"] + " <br/><br/>" + ConfigurationManager.AppSettings["ResponseMessage5"] + " <br/> " + ConfigurationManager.AppSettings["ResponseMessage6"] + " <br/><p>&nbsp;</p></body></html>";
                            objEmail.AutoResponseID = Convert.ToInt64(dr["ID"]);
                            wslog.WriteLog("INF", "Started Sending acknowledgement mail from(" + objEmail.MailFrom + ") to(" + objEmail.MailTo + ") with autoresponse code:" + objEmail.AutoResponseCode);
                            objEmailPusher.SendEmail(objEmail, "");
                            wslog.WriteLog("INF", "Successfully Completed Sending acknowledgement mail from(" + objEmail.MailFrom + ") to(" + objEmail.MailTo + ") with autoresponse code:" + objEmail.AutoResponseCode);
                            objEmailRepository.UpdateSingleAutoResponseRecord(objEmail.AutoResponseID);
                        }
                        catch (Exception ex)
                        {
                            wslog.WriteLog("ERR", ex.Message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                wslog.WriteLog("ERR", ex.Message);
            }
        }

        


        private string ExtractRecentMailContent(string htmlContent,string threadID="")
        {            
            string source = WebUtility.HtmlDecode(htmlContent);
            HtmlDocument resultat = new HtmlDocument();
            resultat.OptionUseIdAttribute = true;
            resultat.LoadHtml(source);            
            //resultat.

            string oldContent = string.Empty;
            var resultatCopy = resultat.DocumentNode;

            List<HtmlNode> divOldContent = resultatCopy.Descendants().Where(x => x.Id == "Lisa3Content").ToList();
            foreach (HtmlNode item in divOldContent)
            {                
                item.Remove();
                HtmlNode bodyNode = resultatCopy.Descendants().FirstOrDefault(x => x.Name.ToUpper() == "BODY");
                if (bodyNode != null)
                {
                    List<HtmlNode> paraList=bodyNode.Descendants().Where(x => x.Name == "p").ToList();
                    foreach (HtmlNode para in paraList)
                    {
                        if (para.InnerHtml.Contains("From:") && para.InnerHtml.Contains("To:") && para.InnerHtml.Contains("Subject:"))
                        {
                            para.Remove();
                            break;
                        }
                    }
                    resultat.OptionUseIdAttribute = true;
                    HtmlNode hNode = resultat.CreateElement("div");
                    hNode.Attributes.Add("id", "LisaMessageThreadIdentifier");
                    hNode.Attributes.Add("style", "visibility:hidden");
                    //hNode.Id = "LisaMessageThreadIdentifier";                    
                    hNode.InnerHtml = threadID;                    
                    bodyNode.AppendChild(hNode);                    
                }
            }                        

            htmlContent = resultat.DocumentNode.OuterHtml;

            return htmlContent;
        }

        static byte[] GetBytes(string str)
        {
            byte[] bytes = new byte[str.Length * sizeof(char)];
            System.Buffer.BlockCopy(str.ToCharArray(), 0, bytes, 0, bytes.Length);
            return bytes;
        }

        static string GetString(byte[] bytes)
        {
            char[] chars = new char[bytes.Length / sizeof(char)];
            System.Buffer.BlockCopy(bytes, 0, chars, 0, bytes.Length);
            return new string(chars);
        }

        public void ExecuteScheduledJob(string ProcedureName, int refreshType = 0)
        {
            var wslog = new Log();
            wslog.WriteLog("INF", "Function call for executing scheduled job explicitly on database");
            try
            {
                var emailRepository = new EmailRepository();
                emailRepository.ExecuteScheduledJob(ProcedureName, refreshType);
            }
            catch (Exception ex)
            {
                wslog.WriteLog("ERR", ex.Message);
            }
        }

        public static bool CheckExcludeAttachmentTypes(string AttachmentExtention)
        {
            bool flag = true;
            string AttachmentType = System.Configuration.ConfigurationManager.AppSettings["AttachmentType"];
            if (!string.IsNullOrEmpty(AttachmentType))
            {
                var myList = AttachmentType.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries).ToList<string>();
                if (myList.Contains(AttachmentExtention.ToUpper()))
                {
                    flag = false;
                }
            }
            return flag;
        }

    }

}
