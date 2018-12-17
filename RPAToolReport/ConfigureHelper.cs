using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAToolReport
{
    public class ConfigureHelper
    {
        public static string GetAppSettingsKeyValue(string keyName)
        {
            return System.Configuration.ConfigurationSettings.AppSettings[keyName];
        }
    }
}
