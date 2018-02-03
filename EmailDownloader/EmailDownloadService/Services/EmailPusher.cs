using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using Microsoft.Exchange.WebServices.Data;
using EmailDownloadService.Models;
using EmailDownloadService.Services;
using System.Net.Mail;
using System.Configuration;
using System.Net.Mime;
using System.Net;
using HtmlAgilityPack;

namespace EmailDownloadService.Repository
{
    public class EmailPusher
    {


        public void SendEmail(Email _ObjEmail, string AttachmentPath)
        {
            string Username, Password, FromEmailAdress, OutHost, outEmailProtocol;
            int OutPort = 0;
            EmailServices emailService = new EmailServices();
            var _listFromEmail = emailService.GetCredentials().FirstOrDefault(x => x.SmtpFromAddress == _ObjEmail.MailFrom);
            if (_listFromEmail != null)
            {
                Username = _listFromEmail.SmtpUser;
                Password = _listFromEmail.SmtpPassword;
                FromEmailAdress = _listFromEmail.SmtpFromAddress;
                outEmailProtocol = _listFromEmail.OutEmailProtocol;
                OutHost = _listFromEmail.OutHost;
                OutPort = _listFromEmail.OutPort;

                if (outEmailProtocol == "Exchange") // Added By Bidhan EmailProtocol == "POP3", will removed on production.
                {
                    //Email _ObjEmail = new Email();
                    //DispositionAction dispositionAction = new DispositionAction();
                    const decimal kilobyte = 1024;
                    ExchangeService service1 = new ExchangeService(ExchangeVersion.Exchange2010_SP1)
                    {
                        Credentials = new WebCredentials(Username, Password)
                    };
                    //to add logic for itemview
                    service1.AutodiscoverUrl(FromEmailAdress, RedirectionUrlValidationCallback);
                    EmailMessage emailmessage = new EmailMessage(service1);
                    emailmessage.Body = _ObjEmail.MailBody;
                    if (_ObjEmail.MailTo != null && _ObjEmail.MailTo != "")
                    {
                        foreach (string emailTo in _ObjEmail.MailTo.Split(';'))
                        {
                            emailmessage.ToRecipients.Add(emailTo);
                        }
                    }
                    if (_ObjEmail.MailSubject != null && _ObjEmail.MailSubject != "")
                        emailmessage.Subject = Convert.ToString(_ObjEmail.MailSubject);
                    if (_ObjEmail.MailCC != null && _ObjEmail.MailCC != "")
                    {
                        foreach (string emailCC in _ObjEmail.MailCC.Split(';'))
                        {
                            emailmessage.CcRecipients.Add(emailCC);
                        }
                    }

                    if (AttachmentPath.Length > 0)
                    {
                        emailmessage.Attachments.AddFileAttachment(AttachmentPath);
                    }

                    emailmessage.Send();
                }
                else if (outEmailProtocol.ToUpper() == "SMTP")
                {

                    const decimal kilobyte = 1024;

                    MailMessage message = new MailMessage();
                    if (_ObjEmail.MailTo != null && _ObjEmail.MailTo != "")
                    {
                        foreach (string emailTo in _ObjEmail.MailTo.Split(';'))
                        {
                            message.To.Add(emailTo);
                        }
                    }
                    if (_ObjEmail.MailSubject != null && _ObjEmail.MailSubject != "")
                        message.Subject = Convert.ToString(_ObjEmail.MailSubject);
                    if (_ObjEmail.MailCC != null && _ObjEmail.MailCC != "")
                    {
                        foreach (string emailCC in _ObjEmail.MailCC.Split(';'))
                        {
                            message.CC.Add(emailCC);
                        }
                    }

                    string strMailBody = _ObjEmail.MailBody;
                    var view = AlternateView.CreateAlternateViewFromString(strMailBody, null, "text/html");
                    message.AlternateViews.Add(view);
                    message.From = new MailAddress(FromEmailAdress);
                    message.IsBodyHtml = true;

                    System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient(OutHost);
                    System.Net.NetworkCredential smtpAuthen = new System.Net.NetworkCredential(Username, Password);
                    smtp.UseDefaultCredentials = false;
                    smtp.Credentials = smtpAuthen;

                    if (AttachmentPath.Length > 0)
                    {

                        message.Attachments.Add(new System.Net.Mail.Attachment(AttachmentPath));
                    }

                    smtp.Send(message);
                }
            }
        }
        
        
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

        private EmailMessage AddInlineImagesExchange(EmailMessage emailmessage, string disResponse, ref string resMessage)
        {

            string appImageRelPath = ConfigurationManager.AppSettings["ApplicationImagePath"];
            string appImageAbsPath = ConfigurationManager.AppSettings["AppImageAbsolutePath"];

            int index = 0;

            while (disResponse.IndexOf("src=\"" + appImageRelPath, 0) > 0)
            {
                int startIndex = disResponse.IndexOf("src=\"" + appImageRelPath);
                int endIndex1 = disResponse.IndexOf("StartLisacid_", 0);
                int endIndex2 = disResponse.IndexOf("_EndLisacid_", 0);
                int endIndex3 = disResponse.IndexOf("_EndLisaImage", 0);
                if (endIndex1 > startIndex && endIndex2 > startIndex && endIndex3 > startIndex)
                {
                    string repContent1 = disResponse.Substring(startIndex, endIndex1 - startIndex);
                    string repContent2 = disResponse.Substring(startIndex, endIndex2 - startIndex);
                    string repContent3 = disResponse.Substring(startIndex, endIndex3 - startIndex);

                    string imageFileName = repContent3.Replace(repContent2 + "_EndLisacid_", "").Replace('_', '.');
                    string imageFilePath = repContent3.Replace("src=\"", "") + "_EndLisaImage" + imageFileName;
                    imageFilePath = imageFilePath.Replace(appImageRelPath, appImageAbsPath);
                    //string contentID = repContent2.Replace(repContent1, "").Replace("StartLisacid_", "").Replace('_', '.');
                    string contentID = Guid.NewGuid().ToString() + imageFileName;
                    disResponse = disResponse.Replace(repContent3 + "_EndLisaImage" + imageFileName, "src=\"cid:" + contentID);

                    emailmessage.Attachments.AddFileAttachment(contentID, imageFilePath);
                    emailmessage.Attachments[index].IsInline = true;
                    emailmessage.Attachments[index].ContentId = contentID;

                    index++;
                }
                else
                {
                    break;
                }
            }
            resMessage = disResponse;
            emailmessage.Body = disResponse;
            return emailmessage;
        }

        private List<LinkedResource> AddInlineImagesSMTP(MailMessage mailmessage, string disResponse, ref string resMessage)
        {

            string appImageRelPath = ConfigurationManager.AppSettings["ApplicationImagePath"];
            string appImageAbsPath = ConfigurationManager.AppSettings["ApplicationImagePath"];
            string appImageUploadedPath = ConfigurationManager.AppSettings["ApplicationImageUploadedPath"];

            int index = 0;
            List<LinkedResource> listLR = new List<LinkedResource>();

            while (disResponse.IndexOf("src=\"" + appImageUploadedPath, 0) > 0)
            {
                int startIndex = disResponse.IndexOf("src=\"" + appImageUploadedPath);
                int endIndex1 = disResponse.IndexOf("StartLisacid_", 0);
                int endIndex2 = disResponse.IndexOf("_EndLisacid_", 0);
                int endIndex3 = disResponse.IndexOf("_EndLisaImage", 0);
                if (endIndex1 > startIndex && endIndex2 > startIndex && endIndex3 > startIndex)
                {
                    string repContent1 = disResponse.Substring(startIndex, endIndex1 - startIndex);
                    string repContent2 = disResponse.Substring(startIndex, endIndex2 - startIndex);
                    string repContent3 = disResponse.Substring(startIndex, endIndex3 - startIndex);

                    string imageFileName = repContent3.Replace(repContent2 + "_EndLisacid_", "").Replace('_', '.');
                    string imageFilePath = repContent3.Replace("src=\"", "") + "_EndLisaImage" + imageFileName;
                    imageFilePath = imageFilePath.Replace(appImageUploadedPath, appImageAbsPath + "\\UploadedImages\\");
                    string contentID = Guid.NewGuid().ToString() + imageFileName;
                    //var inlineImage = new LinkedResource(HttpContext.Current.Server.MapPath(imageFilePath));
                    var inlineImage = new LinkedResource(imageFilePath);
                    inlineImage.ContentId = contentID;
                    listLR.Add(inlineImage);
                    //string contentID = repContent2.Replace(repContent1, "").Replace("StartLisacid_", "").Replace('_', '.');

                    disResponse = disResponse.Replace(repContent3 + "_EndLisaImage" + imageFileName, "src=\"cid:" + contentID);

                    index++;
                }
                else
                {
                    break;
                }
            }


            while (disResponse.IndexOf("src=\"" + appImageRelPath, 0) > 0)
            {
                int startIndex = disResponse.IndexOf("src=\"" + appImageRelPath);
                int endIndex1 = disResponse.IndexOf("StartLisacid_", 0);
                int endIndex2 = disResponse.IndexOf("_EndLisacid_", 0);
                int endIndex3 = disResponse.IndexOf("_EndLisaImage", 0);
                if (endIndex1 > startIndex && endIndex2 > startIndex && endIndex3 > startIndex)
                {
                    string repContent1 = disResponse.Substring(startIndex, endIndex1 - startIndex);
                    string repContent2 = disResponse.Substring(startIndex, endIndex2 - startIndex);
                    string repContent3 = disResponse.Substring(startIndex, endIndex3 - startIndex);

                    string imageFileName = repContent3.Replace(repContent2 + "_EndLisacid_", "").Replace('_', '.');
                    string imageFilePath = repContent3.Replace("src=\"", "") + "_EndLisaImage" + imageFileName;
                    imageFilePath = imageFilePath.Replace(appImageRelPath, appImageAbsPath);
                    string contentID = Guid.NewGuid().ToString() + imageFileName;
                    //var inlineImage = new LinkedResource(HttpContext.Current.Server.MapPath(imageFilePath));
                    var inlineImage = new LinkedResource(imageFilePath);
                    inlineImage.ContentId = contentID;
                    listLR.Add(inlineImage);
                    //string contentID = repContent2.Replace(repContent1, "").Replace("StartLisacid_", "").Replace('_', '.');
                    
                    disResponse = disResponse.Replace(repContent3 + "_EndLisaImage" + imageFileName, "src=\"cid:" + contentID);

                    index++;
                }
                else
                {
                    break;
                }
            }
            
            resMessage = disResponse;
                 return listLR;
        }

        private string WrapMailContent(string htmlContent)
        {
            string source = WebUtility.HtmlDecode(htmlContent);
            HtmlDocument resultat = new HtmlDocument();
            resultat.LoadHtml(source);

            string oldContent = string.Empty;
            var resultatCopy = resultat.DocumentNode;

            HtmlNode htmlNodes = resultatCopy.Descendants().FirstOrDefault(x => x.Name.ToUpper() == "BODY");
            if (htmlNodes != null)
            {
                HtmlNode hNode = resultat.CreateElement("div");
                hNode.Id = "Lisa3Content";
                htmlNodes.PrependChild(hNode);
            }
            else
            {
                htmlContent = "<div id=\"Lisa3Content\">" + resultat.DocumentNode.OuterHtml + "</div>";
            }

            //htmlContent = resultat.DocumentNode.OuterHtml;

            return htmlContent;
        }

        private string AddFormattedHeader(Email _ObjEmail, string fromAddress)
        {
            string toEmailIDs = string.Empty;
            string ccEmailIDs = string.Empty;
            string subjectLine = string.Empty;
            string strHeader = "<div style=\"border:none;border-top:solid #B5C4DF 1.0pt;padding:3.0pt 0in 0in 0in\">";

            if (_ObjEmail.MailTo != null && _ObjEmail.MailTo != "")
            {
                foreach (string emailTo in _ObjEmail.MailTo.Split(','))
                {
                    toEmailIDs = toEmailIDs + emailTo + ";";
                }
                if (toEmailIDs.EndsWith(";"))
                    toEmailIDs = toEmailIDs.Substring(0, toEmailIDs.Length - 1);
            }
            if (_ObjEmail.MailSubject != null && _ObjEmail.MailSubject != "")
                subjectLine = _ObjEmail.MailSubject;
            if (_ObjEmail.MailCC != null && _ObjEmail.MailCC != "")
            {
                foreach (string emailCC in _ObjEmail.MailCC.Split(','))
                {
                    ccEmailIDs = ccEmailIDs + emailCC + ";";
                }
                if (ccEmailIDs.EndsWith(";"))
                    ccEmailIDs = ccEmailIDs.Substring(0, ccEmailIDs.Length - 1);
            }
            strHeader = strHeader + "<p class=\"MsoNormal\"><b><span style=\"font-size:10.0pt;font-family:\" tahoma\",\"sans-serif\"\"=\"\">From:</span></b><span style=\"font-size:10.0pt;font-family:\" tahoma\",\"sans-serif\"\"=\"\"> " + fromAddress + "<br>";
            strHeader = strHeader + "<b>Sent:</b>" + String.Format("{0:F}", DateTime.Now) + "<br>";
            if (toEmailIDs.Length > 0)
            {
                strHeader = strHeader + "<b>To:</b>" + toEmailIDs + "<br>";
            }
            if (ccEmailIDs.Length > 0)
            {
                strHeader = strHeader + "<b>Cc:</b>" + ccEmailIDs + "<br>";
            }

            if (subjectLine.Length > 0)
            {
                strHeader = strHeader + "<b>Subject:</b>" + subjectLine;
            }
            strHeader = strHeader + "</span></p></div>";

            return strHeader;

        }


    }
}