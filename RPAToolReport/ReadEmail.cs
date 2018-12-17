using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RPAToolReport
{
    public class ReadEmail
    {
        private Outlook.MAPIFolder mainbox = null;
        private Outlook.MailItem item = null;
        private AccessDBConnection _accessDBConnection;

        public ReadEmail()
        {
            _accessDBConnection = new AccessDBConnection();
            _accessDBConnection.InitialDbConnection();
        }
        public void readEmailByFolderAndSaveToDb()
        {
            Outlook.Application application = new Outlook.Application();
            Outlook.NameSpace ns = application.GetNamespace("mapi");
            mainbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.MAPIFolder targetFolder = mainbox.Folders[ConfigureHelper.GetAppSettingsKeyValue(Constants.FolderA)];
            Outlook.MAPIFolder processFolder = mainbox.Folders[ConfigureHelper.GetAppSettingsKeyValue(Constants.ProcessFolder)];

            int loopTime = targetFolder.Items.Count;
            for (int i = 0; i < loopTime; i++)
            {
                item = targetFolder.Items.GetFirst();
                SaveEmailInfo(item);
                targetFolder.Items.GetFirst().Move(processFolder);
            }
            _accessDBConnection.CloseDbConnection();
        }
        public void SaveEmailInfo(Outlook.MailItem mailItem)
        {
            if (mailItem.Subject == "Server Down")
            {
                _accessDBConnection.InsertEmailinfo(mailItem);
            }
        }
    }
}
