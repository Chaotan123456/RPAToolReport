using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAToolReport
{
    public class GenerateSQl
    {
        public string insertSqlCommend(string to,string cc,string subject,string body)
        {
            string insert = String.Format("INSERT INTO T_EmailInfo(To,CC,Subject,Body) values('{0}','{1}','{2}','{3}')", to, cc,subject, body);
            return insert;
        }
    }
}
