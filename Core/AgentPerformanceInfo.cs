using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dax.Scrapping.Appraisal.Core
{
    public class AgentPerformanceInfo
    {
        public string Rank { get; set; }
        public string TimeUtilization { get; set; }
        public string EstWorkingTime { get; set; }
        public string TotalLPH { get; set; }
        public string TotalLeads { get; set; }
        public string TotalNoXFER { get; set; }
        public string TotalProjectedLead { get; set; }
    }
}
