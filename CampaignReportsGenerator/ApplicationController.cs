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
        private Mutex _queueMutex = new Mutex();
        private Queue<CampaignReports> _campaignReportsQueue = new Queue<CampaignReports>();
        private Thread pollingThreadStaging = null;
        private Thread pollingThreadProduction = null;
        private Thread processThread = null;

        public ApplicationController()
        {
            this.LoadConfig();
            SharedClass.Logger.Info("Service started");
        }
        public void Start()
        {
            pollingThreadProduction = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(StartDbPolling));
            pollingThreadProduction.Name = "PollerProduction";
            pollingThreadProduction.Start(SourceDatabase.PRODUCTION);

            if (SharedClass.PollStaging)
            {
                pollingThreadStaging = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(StartDbPolling));
                pollingThreadStaging.Name = "PollerStaging";
                pollingThreadStaging.Start(SourceDatabase.STAGING);
            }

            processThread = new Thread(new ThreadStart(this.ProcessRequest));
            processThread.Name = "ProcessThread";
            processThread.Start();

        }
        private void StartDbPolling(object input)
        {
            try
            {
                SourceDatabase sourceDataBase = (SourceDatabase)input;
                SharedClass.Logger.Info("Preparing Connection & Command Objects");
                SqlConnection sqlCon = new SqlConnection(SharedClass.GetConnectionString(sourceDataBase));
                SqlCommand sqlCmd = new SqlCommand("GetPendingCampaignsForReportsGeneration", sqlCon);
                DataSet ds = null;
                SqlDataAdapter da = null;
                byte lastFetchCount = 0;
                sqlCmd.CommandType = CommandType.StoredProcedure;
                SharedClass.Logger.Info("Started " + sourceDataBase.ToString());
                while (!SharedClass.HasStopSignal)
                {
                    try
                    {
                        this._isIamPolling = true;
                        da = new SqlDataAdapter(sqlCmd);
                        ds = new DataSet();
                        da.Fill(ds);
                        lastFetchCount = 0;
                        if (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
                        {
                            lastFetchCount = Convert.ToByte(ds.Tables[0].Rows.Count);
                            foreach (DataRow dr in ds.Tables[0].Rows)
                            {
                                CampaignReports _campaignReportsObj = new CampaignReports();
                                _campaignReportsObj.RequestId = Convert.ToInt64(dr["Id"].ToString());
                                if (!dr["CampaignScheduleId"].Equals(System.DBNull.Value))
                                    _campaignReportsObj.CampaignScheduleId = Convert.ToInt64(dr["CampaignScheduleId"].ToString());
                                else
                                    _campaignReportsObj.CampaignScheduleId = 0;
                                _campaignReportsObj.SourceDataBase = sourceDataBase;
                                
                                this.Enqueue(_campaignReportsObj);

                            }
                        }
                    }
                    catch (Exception e)
                    {
                        SharedClass.Logger.Error(e.ToString());
                    }
                    finally
                    {
                        this._isIamPolling = false;
                    }
                    if (lastFetchCount == 0)
                    {
                        Thread.Sleep(SharedClass.PollingInterval * 1000);
                    }
                }
            }
            catch (Exception ex)
            {
                SharedClass.Logger.Error("Exception in StartDBPool Reason : " + ex.ToString());
            }

        }

        private void Enqueue(CampaignReports _campaignReportsObj)
        {
            try
            {
                while (!this._queueMutex.WaitOne())
                    Thread.Sleep(100);
                SharedClass.Logger.Info("Enqueuing the Request of ID: " + _campaignReportsObj.RequestId.ToString());
                this._campaignReportsQueue.Enqueue(_campaignReportsObj);
            }
            catch (Exception ex)
            {
                SharedClass.Logger.Error("Exception in Enqueuing ReportsObj  Reason : " + ex.ToString());
            }
            finally
            {
                this._queueMutex.ReleaseMutex();
            }
        }

        private int QueueCount()
        {
            int count = 0;
            try
            {
                while (!this._queueMutex.WaitOne())
                    Thread.Sleep(100);
                count = this._campaignReportsQueue.Count();

            }
            catch (Exception ex)
            {
                count = 0;
                SharedClass.Logger.Error("Exception in getting queue count Reason : " + ex.ToString());
            }
            finally
            {
                this._queueMutex.ReleaseMutex();
            }
            return count;
        }

        private CampaignReports DeQueue()
        {
            CampaignReports campaignReports = null;
            try
            {
                while (!this._queueMutex.WaitOne())
                    Thread.Sleep(10);
                campaignReports = this._campaignReportsQueue.Dequeue();
            }
            catch (Exception ex)
            {
                SharedClass.Logger.Error("Exception in Dequeuing Reason:" + ex.ToString());
            }
            finally
            {
                this._queueMutex.ReleaseMutex();
            }
            return campaignReports;
        }

        private void ProcessRequest()
        {
            while (!SharedClass.HasStopSignal)
            {
                GenerateExcel generateExcelObj = null;
                try
                {
                    if (this.QueueCount() == 0)
                    {
                        try
                        {
                            Thread.Sleep(2000);
                        }
                        catch (ThreadAbortException ex) { }
                        catch (ThreadInterruptedException ex) { }
                    }
                    else
                    {
                        CampaignReports _campaignReports = this.DeQueue();
                        if (_campaignReports != null)
                        {
                            this._isIamProcessing = true;
                            generateExcelObj = new GenerateExcel();
                            SharedClass.Logger.Info("Processing request of ID : " + _campaignReports.RequestId.ToString());
                            generateExcelObj.GenerateReports(_campaignReports);
                            this._isIamProcessing = false;
                        }
                    }
                }
                catch (Exception ex)
                {
                    SharedClass.Logger.Error("Exception in Processing the request Reason: " + ex.ToString());
                }
                finally
                {
                    generateExcelObj = null;
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

            if (System.Configuration.ConfigurationManager.ConnectionStrings["ConnectionStringStaging"] != null)
            {
                try
                {
                    SharedClass.ConnectionStringStaging = System.Configuration.ConfigurationManager.ConnectionStrings["ConnectionStringStaging"].ConnectionString;
                    SharedClass.PollStaging = true;
                }
                catch (Exception ex)
                {
                    SharedClass.Logger.Error(ex.ToString());
                }
            }

            if (System.Configuration.ConfigurationManager.AppSettings["PollingInterval"] != null)
            {
                byte tempValue = SharedClass.PollingInterval;
                if (byte.TryParse(System.Configuration.ConfigurationManager.AppSettings["PollingInterval"].ToString(), out tempValue))
                    SharedClass.PollingInterval = tempValue;
            }
        }



        public void Stop()
        {
            SharedClass.Logger.Info("Received stop signal");
            SharedClass.HasStopSignal = true;
                while (pollingThreadProduction.ThreadState == ThreadState.Running)
                {
                    SharedClass.Logger.Info("DBPoll thread is still running. ThreadState :  " + this.pollingThreadProduction.ThreadState.ToString());
                    if (this.pollingThreadProduction.ThreadState == ThreadState.WaitSleepJoin)
                    {
                        try
                        {
                            this.pollingThreadProduction.Interrupt();
                        }
                        catch(Exception ex)
                        {
                            SharedClass.Logger.Error("Exception in interupting the DBPoll Reason : " + ex.ToString());
                        }
                    }
                        
                    Thread.Sleep(200);
                }
                if(SharedClass.PollStaging)
                {
                    while(this.pollingThreadStaging.ThreadState == ThreadState.Running)
                    {
                        SharedClass.Logger.Info("DBPollStaging thread is still running. ThreadState :  " + this.pollingThreadStaging.ThreadState.ToString());
                            if(this.pollingThreadStaging.ThreadState == ThreadState.WaitSleepJoin)
                            {
                                try
                                {
                                    this.pollingThreadStaging.Interrupt();
                                }
                                catch(Exception ex)
                                {
                                    SharedClass.Logger.Error("Exception in interupting the DBPollStaging Reason : " + ex.ToString());
                                }
                            }
                                
                        Thread.Sleep(200);
                    }
                }

                while(this._isIamProcessing)
                {
                    SharedClass.Logger.Info("Processing thread is still running , thread state : " + this.processThread.ThreadState.ToString());
                    if(this.processThread.ThreadState == ThreadState.WaitSleepJoin)
                    {
                        try
                        {
                            this.processThread.Interrupt();
                        }
                        catch(Exception ex)
                        {
                            SharedClass.Logger.Error("Exception in interupting the ProcessThread Reason : " + ex.ToString());
                        }
                    }
                    Thread.Sleep(200);
                }

            SharedClass.Logger.Info("Service has stopped");

        }
    }
}
