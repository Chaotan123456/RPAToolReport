using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAToolReport
{
    class Program
    {
        static void Main(string[] args)
        {
            ReadEmail re = new ReadEmail();
            re.readEmailByFolderAndSaveToDb();
            SaveExcel se = new SaveExcel();
            se.GetDataFromAccessDb();
        }
    }
}
