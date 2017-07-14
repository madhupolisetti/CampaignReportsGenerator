using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CampaignReportsGenerator
{
    public class CampaignReports
    {
        private long _campaignScheduleId = 0;
        private long _requestId = 0;
        private SourceDatabase _sourceDataBase = SourceDatabase.PRODUCTION;
        private CampaignType _campaignType = CampaignType.SMS;

        public long CampaignScheduleId
        {
            get { return _campaignScheduleId; }
            set { _campaignScheduleId = value; }
        }
        public long RequestId
        {
            get { return _requestId; }
            set { _requestId = value; }
        }

        

        public SourceDatabase SourceDataBase
        {
            get { return _sourceDataBase; }
            set { _sourceDataBase = value;}

        }

        public CampaignType CType
        {
            get { return this._campaignType; }
            set { _campaignType = value; }
        }

    }
}
