using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;

namespace CampaignReportsGenerator
{
    static class SharedClass
    {
        private static string _connectionString = string.Empty;
        private static byte _pollingInterval = 10;
        private static ILog _logger = null;
        private static bool _hasStopSignal = false;

        public static void InitializeLogger()
        {
            GlobalContext.Properties["LogName"] = DateTime.Now.ToString("yyyyMMdd");
            log4net.Config.XmlConfigurator.Configure();
            _logger = LogManager.GetLogger("Log");
        }

        public static string ConnectionString
        {
            get { return _connectionString; }
            set { _connectionString = value; }
        }
        public static byte PollingInterval
        {
            get { return _pollingInterval; }
            set { _pollingInterval = value; }
        }
        public static ILog Logger
        {
            get
            {
                if (_logger == null)
                {
                    InitializeLogger();
                }
                return _logger;
            }
        }
        public static bool HasStopSignal
        {
            get { return _hasStopSignal; }
            set { _hasStopSignal = value; }
        }
    }
}
