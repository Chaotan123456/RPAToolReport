using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RPAToolReport
{
    public class SendEmail
    {
        public void send()
        {
            Console.WriteLine("Create email");
            Outlook.Application application = new Outlook.Application();
            Outlook.MailItem mailItem = application.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.To = "chao.tan@accenture.com";
            mailItem.CC = "chao.tan@accenture.com";
            mailItem.Subject = "Server Down";
            mailItem.HTMLBody = "";
            mailItem.Display(false);
            mailItem.Send();
            Console.WriteLine("Send email successfully");
            application = null;
            mailItem = null;
        }
    }
}
