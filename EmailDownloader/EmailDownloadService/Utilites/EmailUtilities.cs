using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EmailDownloadService.Utilites
{
    public class EmailUtilities
    {
        string cstrstatus;
        string cstrTitle;
        string cstrDescription;
        string cstrHelpMessage;
        string cstrLocation;
        string cstrMessageType;
        string cstrMessageCode;
        System.Data.DataSet cdsResult;
        System.Data.DataTable cdtResult;
        Int64 RequestID;
        string Returnparam;
        string result;
        private string _ExceptionMessage;
        public string _StackTrace;
        public string _Source;

        public string OperationStatus
        {
            get { return cstrstatus; }
            set { cstrstatus = value; }
        }
        public string Title
        {
            get { return cstrTitle; }
            set { cstrTitle = value; }
        }

        public string Description
        {
            get { return cstrDescription; }
            set { cstrDescription = value; }
        }
        public string HelpMessage
        {
            get { return cstrHelpMessage; }
            set { cstrHelpMessage = value; }
        }
        public string Location
        {
            get { return cstrLocation; }
            set { cstrLocation = value; }
        }

        public string MessageType
        {
            get { return cstrMessageType; }
            set { cstrMessageType = value; }
        }


        public string MessageCode
        {
            get { return cstrMessageCode; }
            set { cstrMessageCode = value; }
        }


        public System.Data.DataSet ResultSet
        {
            get { return cdsResult; }
            set { cdsResult = value; }
        }

        public System.Data.DataTable ResultTable
        {
            get { return cdtResult; }
            set { cdtResult = value; }
        }


        public Int64 Request_ID
        {
            get { return RequestID; }
            set { RequestID = value; }
        }


        public string ExceptionMessage
        {
            get { return _ExceptionMessage; }
            set { _ExceptionMessage = value; }
        }


        public string StackTrace
        {
            get { return _StackTrace; }
            set { _StackTrace = value; }
        }


        public string Source
        {
            get { return _Source; }
            set { _Source = value; }
        }


        public string Return_param
        {
            get { return Returnparam; }
            set { Returnparam = value; }
        }


        public string Return_Result
        {
            get { return result; }
            set { result = value; }
        }


        public EmailUtilities()
        {
            this.cstrstatus = "NA";
            this.cstrTitle = "";
            this.cstrDescription = "";
            this.cstrHelpMessage = "";
            this.cstrLocation = "";
            this.cstrMessageCode = "";
            this.cstrMessageType = "";
            this.RequestID = 0;
            cdsResult = new System.Data.DataSet();
            this.result = "";
        }

        public EmailUtilities(string status)
        {
            this.cstrstatus = status;
            this.cstrDescription = "";
            this.cstrHelpMessage = "";
            this.cstrMessageCode = "";
            this.cstrMessageType = "";
            this.RequestID = 0;
            cdsResult = new System.Data.DataSet();
        }
        public void SetMessage(string status, string Title, string Description, string MessageCode, string MessageType, string HelpMessage)
        {
            cstrstatus = status;
            cstrTitle = Title;
            cstrDescription = Description;
            cstrMessageCode = MessageCode;
            cstrMessageType = MessageType;
            cstrHelpMessage = HelpMessage;
        }

        public void SetMessage(string status, string Title, string Description, string MessageCode, string MessageType, string HelpMessage, System.Data.DataSet ResultSet)
        {
            cstrstatus = status;
            cstrTitle = Title;
            cstrDescription = Description;
            cstrMessageCode = MessageCode;
            cstrMessageType = MessageType;
            cstrHelpMessage = HelpMessage;
            cdsResult = ResultSet;
        }

        public void SetMessage(string status, string Title, string Description, string MessageCode, string MessageType, string HelpMessage, System.Data.DataTable ResultSet)
        {
            cstrstatus = status;
            cstrTitle = Title;
            cstrDescription = Description;
            cstrMessageCode = MessageCode;
            cstrMessageType = MessageType;
            cstrHelpMessage = HelpMessage;
            cdtResult = ResultSet;
        }

        public void SetMessage(string status, string Title, string Description, string MessageCode, string MessageType, string HelpMessage, string sqlresult)
        {
            cstrstatus = status;
            cstrTitle = Title;
            cstrDescription = Description;
            cstrMessageCode = MessageCode;
            cstrMessageType = MessageType;
            cstrHelpMessage = HelpMessage;
            result = sqlresult;
        }
        public void SetMessage(string status, string Title, string Description, Int64 ReqID)
        {
            cstrstatus = status;
            cstrTitle = Title;
            cstrDescription = Description;
            RequestID = ReqID;
        }

        public void SetMessage(string status, string Title, string Description, string retparam)
        {
            cstrstatus = status;
            cstrTitle = Title;
            cstrDescription = Description;
            Returnparam = retparam;
        }
        public void SetMessage(string status, string Title, string Description, string Location, string retparam)
        {
            cstrstatus = status;
            cstrTitle = Title;
            cstrDescription = Description;
            cstrLocation = Location;
            Returnparam = retparam;
        }
        public void SetResult(string sqlresult)
        {
            result = sqlresult;
        }
    }
}
