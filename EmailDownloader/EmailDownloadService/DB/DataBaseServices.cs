using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using EmailDownloadService.Logs;
using EmailDownloadService.Utilites;

namespace EmailDownloadService.DB
{
    public class DataBaseServices
    {
        Log _wslog;
        public static string MstrConnectionString;

        public DataBaseServices()
        {
            MstrConnectionString = ConfigurationManager.ConnectionStrings["EmailCS"].ConnectionString;
        }
        public SqlConnection ATS_getconnection() 
        {
            using (var con = new SqlConnection(MstrConnectionString))
            {
                con.Open();
                return con;
            }
        }
        public EmailUtilities OpenSQLConnection(SqlConnection sqlCon, bool useExistingConnection)
        {

            EmailUtilities openIrmsStatus = new EmailUtilities();
            try
            {
                if (sqlCon == null)
                    sqlCon = new SqlConnection();
                sqlCon.ConnectionString = MstrConnectionString;
                sqlCon.Open();
                switch (sqlCon.State)
                {
                    case ConnectionState.Broken:
                        sqlCon.Close();
                        sqlCon.Open();
                        break;
                    case ConnectionState.Closed:
                        sqlCon.Open();
                        break;
                    case ConnectionState.Open:
                        if (!useExistingConnection)
                        {
                            sqlCon.Close();
                            sqlCon.Open();
                        }
                        break;
                    default:
                        sqlCon.Open();
                        break;
                }
                openIrmsStatus.SetMessage("Success", "Connection established", "Connection has been established successfully.", "DBA -> DBACCESS", "Establish connection between Business Component and Database Server", "Asked operation completed");
            }
            catch (Exception ex)
            {
                _wslog = new Log();
                var msg = ex.Message + " " + ex.InnerException.Message + " Operation failed : Establish connection between Business Component and Database Server";
                _wslog.WriteLog("ERR", msg);
                openIrmsStatus.SetMessage("Error", ex.Message, ex.InnerException.Message, "DBA -> DBACCESS", "Establish connection between Business Component and Database Server", "Operation failed");
            }
            return openIrmsStatus;
        }
        public EmailUtilities CloseSQLConnection(SqlConnection sqlCon, bool dispose)
        {
            var closeIrmsStatus = new EmailUtilities();
            try
            {
                if (sqlCon.State != ConnectionState.Closed)
                {
                    sqlCon.Close();
                    if (dispose)
                    {
                        sqlCon.Dispose();
                    }
                }
                closeIrmsStatus.SetMessage("Success", "Connection closed", "Connection has been closed successfully.", "DBA -> DBACCESS", "Close established connection between Business Component and Database Server", "Asked operation completed");
            }
            catch (Exception ex)
            {
                _wslog = new Log();
                var msg = ex.Message + " " + ex.InnerException.Message + " Operation failed : Closing establised connection between Business Component and Database Server";
                _wslog.WriteLog("ERR", msg);
                closeIrmsStatus.SetMessage("Error", ex.Message, ex.InnerException.Message, "DBA -> DBACCESS", "Closing establised connection between Business Component and Database Server", "Operation failed");
            }
            return closeIrmsStatus;
        }
        public EmailUtilities PopulateDataSet(string tsqlQuery)
        {
            var popIrmsStatus = new EmailUtilities();
            try
            {
                using (var lDSet = new DataSet())
                {
                    var lSqlCon = new SqlConnection();
                    popIrmsStatus = OpenSQLConnection(lSqlCon, true);
                    var lDa = new SqlDataAdapter(tsqlQuery, lSqlCon) { SelectCommand = { CommandTimeout = 0 } };
                    lDa.Fill(lDSet);
                    lDa.Dispose();
                    popIrmsStatus = CloseSQLConnection(lSqlCon, true);
                    popIrmsStatus.SetMessage("Success", "Popultate Dataset", "Dataset has been populated successfully.", "DBA -> DBACCESS", "Populating SQL query", "Asked operation completed", lDSet);
                }
            }
            catch (Exception exp)
            {
                _wslog = new Log();
                var msg = exp.Message + " " + exp.InnerException.Message;
                _wslog.WriteLog("ERR", msg);
                popIrmsStatus.SetMessage("Error", exp.Source, exp.Message, exp.StackTrace, "");
            }

            return popIrmsStatus;
        }
        public EmailUtilities RetrieveData_by_SQLCommand(SqlCommand cmd)
        {
            var popIrmsStatus = new EmailUtilities();
            try
            {
                using (var lDSet = new DataSet())
                {
                    var lSqlCon = new SqlConnection();
                    popIrmsStatus = OpenSQLConnection(lSqlCon, true);
                    cmd.Connection = lSqlCon;
                    cmd.CommandType = CommandType.StoredProcedure;
                    var lDa = new SqlDataAdapter(cmd) { SelectCommand = { CommandTimeout = 0 } };
                    lDa.Fill(lDSet);
                    lDa.Dispose();
                    popIrmsStatus = CloseSQLConnection(lSqlCon, true);
                    popIrmsStatus.SetMessage("Success", "Popultate Dataset", "Dataset has been populated successfully.", "DBA -> DBACCESS", "Populating SQL query", "Asked operation completed", lDSet);
                }
            }
            catch (Exception exp)
            {
                _wslog = new Log();
                var msg = exp.Message;
                _wslog.WriteLog("ERR", msg);
                popIrmsStatus.SetMessage("Error", exp.Source, exp.Message, exp.StackTrace);
            }
            return popIrmsStatus;
        }

        public EmailUtilities RetrieveDataTable_by_SQLCommand(SqlCommand cmd)
        {
            var popIrmsStatus = new EmailUtilities();
            try
            {
                using (var lDTable = new DataTable())
                {
                    var lSqlCon = new SqlConnection();
                    popIrmsStatus = OpenSQLConnection(lSqlCon, true);
                    cmd.Connection = lSqlCon;
                    cmd.CommandType = CommandType.StoredProcedure;
                    var lDa = new SqlDataAdapter(cmd) { SelectCommand = { CommandTimeout = 0 } };
                    lDa.Fill(lDTable);
                    lDa.Dispose();
                    popIrmsStatus = CloseSQLConnection(lSqlCon, true);
                    popIrmsStatus.SetMessage("Success", "Popultate DataTable", "DataTable has been populated successfully.", "DBA -> DBACCESS", "Populating SQL query", "Asked operation completed", lDTable);
                }
            }
            catch (Exception exp)
            {
                _wslog = new Log();
                var msg = exp.Message;
                _wslog.WriteLog("ERR", msg);
                popIrmsStatus.SetMessage("Error", exp.Source, exp.Message, exp.StackTrace);
            }
            return popIrmsStatus;
        }
       
        public EmailUtilities ExecuteNonQuery(string storedProcedureName, params SqlParameter[] arrParam)
        {
            SqlParameter firstOutputParameter = null;
            SqlParameter charOutputParameter = null;
            Int64 retVal = 0;
            string retvalue = "";
            var objStaus = new EmailUtilities();
            var lSqlCon = new SqlConnection(MstrConnectionString);
            try
            {
                objStaus = OpenSQLConnection(lSqlCon, true);
                long lngRowsaffected;
                SqlCommand lSqlComm;
                using (lSqlComm = new SqlCommand())
                {
                    lSqlComm.Connection = lSqlCon;
                    lSqlComm.CommandText = storedProcedureName;
                    lSqlComm.CommandType = CommandType.StoredProcedure;
                    lSqlComm.CommandTimeout = 0;

                    if (arrParam != null)
                    {
                        foreach (SqlParameter param in arrParam)
                        {
                            lSqlComm.Parameters.Add(param);
                            if (firstOutputParameter == null && param.Direction == ParameterDirection.Output && param.SqlDbType == SqlDbType.BigInt)
                            {
                                firstOutputParameter = param;
                            }

                            if (charOutputParameter == null && param.Direction == ParameterDirection.Output && param.SqlDbType == SqlDbType.VarChar)
                            {
                                charOutputParameter = param;
                            }
                        }
                    }
                    lngRowsaffected = lSqlComm.ExecuteNonQuery();
                    if (firstOutputParameter != null)
                    {
                        retVal = (Int64)firstOutputParameter.Value;
                    }
                    if (charOutputParameter != null)
                    {
                        retvalue = charOutputParameter.Value.ToString();
                    }
                }

                objStaus = CloseSQLConnection(lSqlCon, true);
                objStaus.SetMessage("Success", "Data Updated", "Total no. of Rows affected " + lngRowsaffected,
                    storedProcedureName, retvalue != "" ? retvalue : retVal.ToString());
            }
            catch (Exception exp)
            {
                _wslog.WriteLog("ERR", exp.Message);
                objStaus.SetMessage("Error", exp.Source, exp.Message, exp.StackTrace,
                    retvalue != "" ? retvalue : retVal.ToString());
            }
            return objStaus;
        }

        public EmailUtilities ExecuteNonQuery(string storedProcedureName)
        {
            Int64 retVal = 0;
            string retvalue = "";
            var objStaus = new EmailUtilities();
            var lSqlCon = new SqlConnection(MstrConnectionString);
            try
            {
                objStaus = OpenSQLConnection(lSqlCon, true);
                long lngRowsaffected;
                SqlCommand lSqlComm;
                using (lSqlComm = new SqlCommand())
                {
                    lSqlComm.Connection = lSqlCon;
                    lSqlComm.CommandText = storedProcedureName;
                    lSqlComm.CommandType = CommandType.StoredProcedure;
                    lSqlComm.CommandTimeout = 0;
                    
                    lngRowsaffected = lSqlComm.ExecuteNonQuery();                    
                }

                objStaus = CloseSQLConnection(lSqlCon, true);
                objStaus.SetMessage("Success", storedProcedureName + " : Job Executed Successfully.", "Total no. of Rows affected " + lngRowsaffected,
                    storedProcedureName, retvalue != "" ? retvalue : retVal.ToString());
            }
            catch (Exception exp)
            {
                _wslog.WriteLog("ERR", exp.Message);
                objStaus.SetMessage("Error", exp.Source, exp.Message, exp.StackTrace,
                    retvalue != "" ? retvalue : retVal.ToString());
            }

            return objStaus;
            
        }

        
    }
}
