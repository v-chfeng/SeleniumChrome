using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CrawlPBI.Mode
{
    public class ReportList
    {
        public static string[] GetReports()
        {
            string[] reportArr = System.Configuration.ConfigurationManager.AppSettings.GetValues("ReportsList");
           
            return reportArr;
        }
    }
}
