using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using EmailDownloadService.DB;
using EmailDownloadService.Logs;
using EmailDownloadService.Models;
using EmailDownloadService.Utilites;

namespace EmailDownloadService.Repository
{
    public class EmailRepository
    {
        Log wslog = new Log();
        public EmailUtilities GetEmailInformation(InformationTypeEnum informationType)
        {
            var result = new EmailUtilities();
            var sqlParametersList = new List<SqlParameter>
            {
                new SqlParameter("@informationType", informationType)
                {
                    SqlDbType = SqlDbType.VarChar
                }
            };
            try
            {
                result = GetData("usp_GetEmailInformation", sqlParametersList);
            }
            catch (Exception ex)
            {
                wslog.WriteLog("ERR", ex.Message.ToString());
                throw new ArgumentException(result.Description, "Error");
            }
            return result;
        }

        public EmailUtilities GetEmailInformation(InformationTypeEnum informationType, string ReceivingEMailID)
        {
            var resultemail = new EmailUtilities();
            var sqlParametersList = new List<SqlParameter>
            {
                new SqlParameter("@informationType", informationType)
                {
                    SqlDbType = SqlDbType.VarChar
                },

                new SqlParameter("@receivingEMailID", ReceivingEMailID)
                {
                    SqlDbType = SqlDbType.VarChar
                }
            };
            try
            {
                resultemail = GetData("usp_GetEmailInformation", sqlParametersList);
            }
            catch (Exception ex)
            {
                wslog.WriteLog("ERR", ex.Message.ToString());
                throw new ArgumentException(resultemail.Description, "Error");
            }
            return resultemail;
        }
        public EmailUtilities InsertPop3MessageGuid(int ID, string ReceivingEMailID = "")
        {
            var resultemaillist = new EmailUtilities();
            var sqlParametersList = new List<SqlParameter>
            {
                new SqlParameter("@uid", ID)
                {
                    SqlDbType = SqlDbType.BigInt
                },
                new SqlParameter("@receivingEMailID", ReceivingEMailID)
                {
                    SqlDbType = SqlDbType.NVarChar
                }
            };
            try
            {
                resultemaillist = SaveData("insertPOP3MessageGUID", sqlParametersList);
            }
            catch (Exception ex)
            {
                wslog.WriteLog("ERR", ex.Message.ToString());
                throw new ArgumentException(resultemaillist.Description, "Error");
            }
            return resultemaillist;
        }
        
        public EmailUtilities SaveEmailDownloadDate(DateTime emailDownloadDate, string smptpFromAddress)
        {
            var resultemailDownload = new EmailUtilities();
            var sqlParametersList = new List<SqlParameter>
            {
                new SqlParameter("@emailDownloadDate", emailDownloadDate)
                {
                    SqlDbType = SqlDbType.DateTime
                },
                new SqlParameter("@SmptpFromAddress", smptpFromAddress)
                {
                    SqlDbType = SqlDbType.NVarChar
                }
            };
            try
            {
                resultemailDownload = SaveData("usp_SaveEmailDownloadDate", sqlParametersList);
            }
            catch (Exception ex)
            {
                wslog.WriteLog("ERR", ex.Message.ToString());
                throw new ArgumentException(resultemailDownload.Description, "Error");
            }
            return resultemailDownload;
        }

        public EmailUtilities UpdatePOP3MessageID(string smptpFromAddress,long messageID)
        {
            var sqlParametersList = new List<SqlParameter>
            {                
                new SqlParameter("@receivingEmailID", smptpFromAddress)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@lastMessageID", messageID)
                {
                    SqlDbType = SqlDbType.BigInt
                }
            };
            return SaveData("usp_UpdatePOP3MessageMaster", sqlParametersList);
        }

        public EmailUtilities DownloadEmail(string flag, string @from, string to, string cc, string subject, string body,string uniqueMessageID,string emailBoxMessageUID,string threadID, DateTime createdOn, string sendername, int messageId,long pop3LastMessageID,string emailCredential,DateTime lastEmailDownloadDate)
        {

            var sqlParametersList = new List<SqlParameter>
            {
                new SqlParameter("@flag", flag)
                {
                    SqlDbType = SqlDbType.VarChar
                },
                new SqlParameter("@From", @from)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@To", to)
                {
                    SqlDbType = SqlDbType.VarChar
                },
                new SqlParameter("@CC", cc)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@Subject", subject)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@Body", body)
                {
                    SqlDbType = SqlDbType.NText
                },
                new SqlParameter("@UniqueMessageID", uniqueMessageID)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@createdOn", createdOn)
                {
                    SqlDbType = SqlDbType.DateTime
                },
                new SqlParameter("@SenderName", sendername)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@messageId",messageId)
                {
                    SqlDbType = SqlDbType.BigInt
                },
                new SqlParameter("@Pop3LastMessageID",pop3LastMessageID)
                {
                    SqlDbType = SqlDbType.BigInt
                },
                new SqlParameter("@emailcredential",emailCredential)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@LastEmailDownloadDate",lastEmailDownloadDate)
                {
                    SqlDbType = SqlDbType.DateTime
                },
                new SqlParameter("@EmailBoxMessageID",emailBoxMessageUID)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@ThreadID",threadID)
                {
                    SqlDbType = SqlDbType.NVarChar
                }
               
            };
            return SaveData("usp_DownloadEmailsToEuroRawData", sqlParametersList);

        }

        public EmailUtilities DownloadAttachments(string attachmentPath, int lisaId, string fileName, decimal fileSize, int MessageId)
        {
            var sqlParametersList = new List<SqlParameter>
            {
                new SqlParameter("@LisaId", lisaId)
                {
                    SqlDbType = SqlDbType.BigInt
                },
                new SqlParameter("@FileName", fileName)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@AttachmentPath", attachmentPath)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@fileSize", fileSize)
                {
                    SqlDbType = SqlDbType.Decimal
                },
                new SqlParameter("@MessageId",MessageId)
                {
                    SqlDbType = SqlDbType.BigInt
               }
                
            };

            return SaveData("usp_saveEmailAttachments", sqlParametersList);
        }

        private static EmailUtilities GetData(string procedureName, List<SqlParameter> parameterList)
        {
            var emailStatusUtility = new EmailUtilities();
            var dbs = new DataBaseServices();
            var wslog = new Log();
            try
            {
                var cmd = new SqlCommand(procedureName) { CommandType = CommandType.StoredProcedure };
                if (parameterList != null)
                    cmd.Parameters.AddRange(parameterList.ToArray());
                emailStatusUtility = dbs.RetrieveDataTable_by_SQLCommand(cmd);
                if (emailStatusUtility.OperationStatus != "Success")
                {
                    wslog.WriteLog("ERR", emailStatusUtility.Description);
                    throw new ArgumentException(emailStatusUtility.Description, "Error");
                }
            }
            catch (Exception exp)
            {
                wslog.WriteLog("ERR", exp.Message.ToString());
            }
            return emailStatusUtility;
        }

        public EmailUtilities SaveData(string procedureName, List<SqlParameter> sqlparameterList)
        {
            var emailStatusUtility = new EmailUtilities();
            var dbs = new DataBaseServices();
            var wslog = new Log();
            try
            {
                emailStatusUtility = dbs.ExecuteNonQuery(procedureName, sqlparameterList.ToArray());
                if (emailStatusUtility.OperationStatus != "Success")
                {
                    wslog.WriteLog("ERR", emailStatusUtility.Description);
                    throw new ArgumentException(emailStatusUtility.Description, "Error");
                }
            }
            catch (Exception exp)
            {
                wslog.WriteLog("ERR", exp.Message.ToString());
            }
            return emailStatusUtility;
        }

        public EmailUtilities GetRemovableText(string ReceivingEmailID)
        {
            var sqlParametersList = new List<SqlParameter>
            {                
                new SqlParameter("@receivingEmailID", ReceivingEmailID)
                {
                    SqlDbType = SqlDbType.NVarChar
                }
            };
            return GetData("[dbo].[GetRemovableText]", sqlParametersList);
        }

        public EmailUtilities GetThreadID(string ReceivingEmailID, string FromEmailID, string ToEmailIDs = "", string CCEmailIDs = "", string Subject = "", string ThreadID = "")
        {
            var sqlParametersList = new List<SqlParameter>
            {                
                new SqlParameter("@ReceivedOnEmail", ReceivingEmailID)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@FromEmail", FromEmailID)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@ToEmail", ToEmailIDs)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@CCEmail", CCEmailIDs)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@ThreadID", ThreadID)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@Subject", Subject)
                {
                    SqlDbType = SqlDbType.NVarChar
                }
            };
            return GetData("[dbo].[usp_CheckThreadExist_Ver1]", sqlParametersList);
        }

        public EmailUtilities AddAutoResponseRecord(string FromMailID, string ReceivingMailID,string Subject, string ThreadID, DateTime MessageCreatedDate, string AutoResponseID)
        {
            var AutoRecord = new EmailUtilities();
            var sqlParametersList = new List<SqlParameter>
            {                
                new SqlParameter("@FromMailID", FromMailID)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@ReceivingMailID", ReceivingMailID)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@Subject", Subject)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@ThreadID", ThreadID)
                {
                    SqlDbType = SqlDbType.NVarChar
                },
                new SqlParameter("@MessageCreatedDate", MessageCreatedDate)
                {
                    SqlDbType = SqlDbType.DateTime
                },
                new SqlParameter("@AutoResponseID", AutoResponseID)
                {
                    SqlDbType = SqlDbType.NVarChar
                }
            };
            try
            {
                AutoRecord = SaveData("[dbo].[usp_AddAutoResponseID]", sqlParametersList);
            }
            catch (Exception exp)
            {
                wslog.WriteLog("ERR", exp.Message.ToString());
                throw new ArgumentException(AutoRecord.Description, "Error");
            }
            return AutoRecord;
        }

        public EmailUtilities GetSingleAutoResponseRecord(string ReceivingMailID="")
        {
            var sqlParametersList = new List<SqlParameter>();
            return GetData("[dbo].[usp_GetSingleAutoResponseID]", sqlParametersList);
        }

        public EmailUtilities GetAllAutoResponseRecord(string ReceivingMailID = "")
        {
            var sqlParametersList = new List<SqlParameter>();
            return GetData("[dbo].[usp_GetAllAutoResponseID]", sqlParametersList);
        }

        public EmailUtilities UpdateSingleAutoResponseRecord(long id)
        {
            var sqlParametersList = new List<SqlParameter>
            {                
                new SqlParameter("@id", id)
                {
                    SqlDbType = SqlDbType.BigInt
                }
            };
            return GetData("[dbo].[usp_UpdateSingleAutoResponseID]", sqlParametersList);
        }
        public void ExecuteScheduledJob(string ProcedureName,int refreshType=0)
        {
            var emailStatusUtility = new EmailUtilities();
            var dbs = new DataBaseServices();
            var wslog = new Log();
            var sqlParametersList = new List<SqlParameter>
            {                
                new SqlParameter("@refreshtype", refreshType)
                {
                    SqlDbType = SqlDbType.NVarChar
                }                
            };
            
            try
            {
                if(refreshType==0)
                    emailStatusUtility = dbs.ExecuteNonQuery(ProcedureName);
                else
                    emailStatusUtility = dbs.ExecuteNonQuery(ProcedureName, sqlParametersList.ToArray());
                if (emailStatusUtility.OperationStatus != "Success")
                {
                    wslog.WriteLog("ERR", emailStatusUtility.Description);
                    throw new ArgumentException(emailStatusUtility.Description, "Error");
                }
            }
            catch (Exception exp)
            {
                wslog.WriteLog("ERR", exp.Message.ToString());
            }
        }

        public EmailUtilities InsertTotalEmailCountOnServer(long TotalEmailCount, long TotalEmailsAndMuidDiff, string ReceivingEMailID)
        {
            var resultemaillist = new EmailUtilities();
            var sqlParametersList = new List<SqlParameter>
            {
                new SqlParameter("@TotalEmailCount", TotalEmailCount)
                {
                    SqlDbType = SqlDbType.BigInt
                },
                new SqlParameter("@TotalEmailsAndMuidDiff", TotalEmailsAndMuidDiff)
                {
                    SqlDbType = SqlDbType.BigInt
                },
                new SqlParameter("@ReceivingEMailID", ReceivingEMailID)
                {
                    SqlDbType = SqlDbType.VarChar
                }
            };
            try
            {
                resultemaillist = SaveData("InsertTotalEmailCountOnServer", sqlParametersList);
            }
            catch (Exception ex)
            {
                wslog.WriteLog("ERR", ex.Message.ToString());
                throw new ArgumentException(resultemaillist.Description, "Error");
            }
            return resultemaillist;
        }
    }
}
