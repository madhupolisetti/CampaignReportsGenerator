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
        private static string _connectionStringStaging = string.Empty;
        private static byte _pollingInterval = 10;
        private static ILog _logger = null;
        private static bool _hasStopSignal = false;
        private static string _savePathProduction = string.Empty;
        private static string _savePathStaging = string.Empty;
        private static bool _pollStaging = false;
        private static Int32 _maxRowsPerSheet = 0;

        public static void InitializeLogger()
        {
            GlobalContext.Properties["LogName"] = DateTime.Now.ToString("yyyyMMdd");
            log4net.Config.XmlConfigurator.Configure();
            _logger = LogManager.GetLogger("Log");
        }

        public static string ConnectionString
        {
            get
            {
                if (_connectionString == null)
                {
                    _connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["DBConnectionString"].ConnectionString;
                }
                return _connectionString;
            }
            set { _connectionString = value; }
        }
        public static string ConnectionStringStaging
        {
            get
            {
                if (_connectionStringStaging == null)
                    _connectionStringStaging = System.Configuration.ConfigurationManager.ConnectionStrings["DBConnectionStringStaging"].ConnectionString;
                return _connectionStringStaging;
            }
            set { _connectionStringStaging = value; }
        }

        public static string GetConnectionString(SourceDatabase sourceDataBase)
        {
            if (sourceDataBase == SourceDatabase.STAGING)
                return SharedClass._connectionStringStaging;
            else
                return SharedClass._connectionString;
        }
        public static string GetFileSavePath(SourceDatabase sourceDataBase)
        {
            if(sourceDataBase == SourceDatabase.STAGING)
            {
                if (_savePathStaging.Length == 0)
                    _savePathStaging = System.Configuration.ConfigurationManager.AppSettings["SavingPathStaging"];
                return _savePathStaging;
            }
            else
            {
                if (_savePathProduction.Length == 0)
                    _savePathProduction = System.Configuration.ConfigurationManager.AppSettings["SavingPathProduction"];
                return _savePathProduction;
            }
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

        public static int MaxRowsPerSheet
        {

            get 
            {
                if(_maxRowsPerSheet == 0)
                {
                    try
                    {
                        _maxRowsPerSheet = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["MaxRowsPerSheet"]);
                    }
                    catch(Exception ex)
                    {
                        SharedClass.Logger.Error("MaxRowsPerSheet key is not set in the config file. Setting to default value 5000");
                        _maxRowsPerSheet = 5000;
                    }
                }
                return _maxRowsPerSheet;
            }
        }

        public static bool PollStaging
        {
            get { return _pollStaging; }
            set { _pollStaging = value; }
        }

        
    }
}
