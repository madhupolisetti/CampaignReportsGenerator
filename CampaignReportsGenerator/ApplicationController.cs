using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Threading;

namespace CampaignReportsGenerator
{
    public class ApplicationController
    {
        private bool _isIamPolling = false;
        private bool _isIamProcessing = false;
        public ApplicationController()
        {
            this.LoadConfig();
        }
        public void Start()
        {
            Thread pollingThread = new Thread(new ThreadStart(this.StartDbPolling));
            pollingThread.Name = "Poller";
            pollingThread.Start();
        }
        private void StartDbPolling()
        {
            SharedClass.Logger.Info("Preparing Connection & Command Objects");
            SqlConnection sqlCon = new SqlConnection(SharedClass.ConnectionString);
            SqlCommand sqlCmd = new SqlCommand("GetPendingCampaignsForReportsGeneration", sqlCon);
            DataSet ds = null;
            SqlDataAdapter da = null;
            byte lastFetchCount = 0;
            sqlCmd.CommandType = CommandType.StoredProcedure;
            SharedClass.Logger.Info("Started");
            while (!SharedClass.HasStopSignal)
            {
                try
                {
                    da = new SqlDataAdapter(sqlCmd);
                    ds = new DataSet();
                    da.Fill(ds);
                    lastFetchCount = 0;
                    if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                    {
                        lastFetchCount = Convert.ToByte(ds.Tables[0].Rows.Count);

                    }
                }
                catch (Exception e)
                {
                    SharedClass.Logger.Error(e.ToString());
                }
                if (lastFetchCount == 0)
                {
                    Thread.Sleep(SharedClass.PollingInterval * 1000);
                }
            }
        }
        private void LoadConfig()
        {
            SharedClass.InitializeLogger();
            try
            {
                SharedClass.ConnectionString = System.Configuration.ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
            }
            catch (Exception e)
            {
                throw new KeyNotFoundException("ConnectionString not found in Application Config", e);
            }
            
            if (System.Configuration.ConfigurationManager.AppSettings["PollingInterval"] != null)
            {
                byte tempValue = SharedClass.PollingInterval;
                if (byte.TryParse(System.Configuration.ConfigurationManager.AppSettings["PollingInterval"].ToString(), out tempValue))
                    SharedClass.PollingInterval = tempValue;
            }
        }
    }
}
