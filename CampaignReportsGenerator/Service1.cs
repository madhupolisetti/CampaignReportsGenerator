using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace CampaignReportsGenerator
{
    public partial class Service1 : ServiceBase
    {
        ApplicationController appController = null;
        System.Threading.Thread appControllerThread = null;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            appController = new ApplicationController();
            appControllerThread = new System.Threading.Thread(new System.Threading.ThreadStart(appController.Start));
            appControllerThread.Name = "ApplicationControllerThread";
            appControllerThread.Start();
 
        }

        protected override void OnStop()
        {
            appController.Stop();
            appControllerThread.Abort();
        }
    }
}
