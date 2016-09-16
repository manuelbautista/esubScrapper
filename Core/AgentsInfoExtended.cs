using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dax.Scrapping.Appraisal.Core
{
    public class AgentsInfoExtended
    {
        public int Id { get; set; }
        public string USER { get; set; }
        public string USERGROUP { get; set; }
        public string Date { get; set; }
        public string CALLS { get; set; }
        public string TIME { get; set; }
        public string PAUSE { get; set; }
        public string PAUSEAVG { get; set; }
        public string WAIT { get; set; }
        public string WAITAVG { get; set; }
        public string TALK { get; set; }
        public string TALKAVG { get; set; }
        public string DISPO { get; set; }
        public string DISPOAVG { get; set; }
        public string DEAD { get; set; }
        public string DEADAVG { get; set; }
        public string CUST { get; set; }
        public string CUSTAVG { get; set; }
        public string X1 { get; set; }
        public string NORSPN { get; set; }
        public string N { get; set; }
        public string NRN { get; set; }
        public string DNQ { get; set; }
        public string INSCH { get; set; }
        public string NOSCH { get; set; }
        public string B { get; set; }
        public string WRONG { get; set; }
        public string CBHOLD { get; set; }
        public string DNC { get; set; }
        public string A { get; set; }
        public string XHT { get; set; }
        public string HS { get; set; }
        public string HU { get; set; }
        public string X2 { get; set; }
        public string CBL { get; set; }
        public string NI { get; set; }
        public string X3 { get; set; }
        public string X4 { get; set; }
        public string JOB { get; set; }
        public string XFER { get; set; }
        public string NOT18 { get; set; }
        public int TotalUnique { get; set; }
        public int TotalSchools { get; set; }
        public double TotalEDS { get; set; }
        public double PercentofTransfers { get; set; }
        public double MatchRate { get; set; }
        public double Hours { get; set; }
        public double UniqueConversionRate { get; set; }
        public double TotalConversionRate { get; set; }
        public double LPH { get; set; }
        public double UPH { get; set; }
        public double CPH { get; set; }
        public string MPT { get; set; }
        public string LogoutTime { get; set; }
        public string RealPause { get; set; }
    }
}
