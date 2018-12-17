using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using Outlook = Microsoft.Office.Interop.Outlook;
namespace RPAToolReport
{
    public class AccessDBConnection
    {
        private static OleDbConnection _dbconn;
        public void InitialDbConnection()
        {

            string Con = @"Provider=Microsoft.Jet.OleDb.4.0;Data Source="+ Environment.CurrentDirectory + System.Configuration.ConfigurationSettings.AppSettings["AccessDBPath"];
            _dbconn = new OleDbConnection(Con);

            _dbconn.Open();
        }

        public void CloseDbConnection()
        {
            _dbconn.Close();
        }

        public void ExcuteQuery(string sql)
        {
            OleDbCommand myCommand = new OleDbCommand(sql, _dbconn);//执行命令
            
            myCommand.ExecuteNonQuery();
        }

        public void InsertEmailinfo(Outlook.MailItem mailItem)
        {
            string sql = @"INSERT INTO T_EmailInfo([To],[CC],[Subject],[Body]) values(@to,@cc,@subject,@body)";
            OleDbCommand myCommand = new OleDbCommand(sql, _dbconn);//执行命令
            myCommand.Parameters.AddWithValue("@to", mailItem.To);
            myCommand.Parameters.AddWithValue("@cc", mailItem.CC);
            myCommand.Parameters.AddWithValue("@subject", mailItem.Subject);
            myCommand.Parameters.AddWithValue("@body", mailItem.Body);
            myCommand.ExecuteNonQuery();
        }

        public DataSet GetDataFromDb(string sql)
        {
            OleDbDataAdapter inst = new OleDbDataAdapter(sql, _dbconn);//选择全部内容
            DataSet ds = new DataSet();//临时存储
            inst.Fill(ds);//用inst填充ds
            return ds;
        }
    }
    
}
