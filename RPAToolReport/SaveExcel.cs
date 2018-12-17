using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAToolReport
{
    public class SaveExcel
    {
        private AccessDBConnection _accessDBConnection;
        public SaveExcel()
        {
            _accessDBConnection = new AccessDBConnection();
            _accessDBConnection.InitialDbConnection();
        }
        public void GetDataFromAccessDb()
        {
            string sql = @"SELECT * FROM T_EmailInfo";
            DataSet ds = _accessDBConnection.GetDataFromDb(sql);
            var dt = ds.Tables[0];
            foreach (DataRow dtRow in dt.Rows)
            {
                string to = dtRow["To"].ToString();
                string cc = dtRow["CC"].ToString();
                string subject = dtRow["Subject"].ToString();
                string body = dtRow["Body"].ToString();
                insertIntoExcel(to,cc,subject,body);
            }
            _accessDBConnection.CloseDbConnection();
        }
        public void insertIntoExcel(string to, string cc,string subject , string body)
        {
            
            OleDbConnectionStringBuilder connectStringBuilder = new OleDbConnectionStringBuilder();
            connectStringBuilder.DataSource = Environment.CurrentDirectory + System.Configuration.ConfigurationSettings.AppSettings["ExcelPath"]; ;
            connectStringBuilder.Provider = "Microsoft.ACE.OLEDB.16.0";
            connectStringBuilder.Add("Extended Properties", "Excel 8.0");
            using (OleDbConnection cn = new OleDbConnection(connectStringBuilder.ConnectionString))
            {
                string sql = "Insert into [Sheet1$] ([To],[CC],[Subject],[Body]) values ('" + to + "','" + cc + "','" + subject + "','" + body + "')";
                OleDbCommand cmdLiming = new OleDbCommand(sql, cn);
                cn.Open();
                cmdLiming.ExecuteNonQuery();
                cn.Close();
            }
        }
    }
}
