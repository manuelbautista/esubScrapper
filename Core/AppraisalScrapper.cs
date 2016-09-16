// Decompiled with JetBrains decompiler
// Type: Dax.Scrapping.Appraisal.Core.AppraisalScrapper
// Assembly: Dax.Scrapping.Appraisal, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 636504A0-34C1-4208-9BC8-AA1181CC609A
// Assembly location: C:\Users\manue\Desktop\EsubScraper\Dax.Scrapping.Appraisal.exe

using CefSharp;
using CefSharp.WinForms;
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Data.Entity.Migrations;
using System.Data.Objects;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Transactions;
using System.Windows.Forms;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace Dax.Scrapping.Appraisal.Core
{
    public class Rank
    {
        public double LPH { get; set; }
        public int UserId { get; set; }
        public string Username { get; set; }
        public double TimeUtilization { get; set; }
    }
  public class AppraisalScrapper : ScrapperBase, IDisposable
  {
    private Action _curAction = (Action) null;
    public Status _curStatus = Status.Paused;
    private string _loginUrl = "https://x5adminw.ytel.com/Account/login";
    public string _newAgentsInfoSite = "https://x5adminw.ytel.com/AgentPerformance/estomescustom";
    private string _user;
    private string _pass;
    private bool isLogged { get; set; }
    public event AppraisalScrapper.LogMsg OnLog = null;

    public event AppraisalScrapper.Completed OnCompleted = null;

    public AppraisalScrapper(string user, string pass)
    {
      this._user = user;
      this._pass = pass;
      this._brouserComponent = new ChromiumWebBrowser(this._loginUrl);
      this._brouserComponent.LoadingStateChanged += new EventHandler<LoadingStateChangedEventArgs>(this._brouserComponent_LoadingStateChanged);
      this._curStatus = Status.Loading;
    }

    private void LoginUser()
    {
      this._brouserComponent.ExecuteScriptAsync("(function() {                                        \n                                    var elemUser = document.getElementById('username');\n                                    elemUser.value = '{0}';\n                                    var elemPass = document.getElementById('password');\n                                    elemPass.value = '{1}';\n                                    document.getElementById('termsckbox').checked = true;\n                                    document.getElementById('login').click();\n                               })();".Replace("{0}", this._user).Replace("{1}", this._pass));
    }

    private void Log(string msg)
    {
      // ISSUE: reference to a compiler-generated field
      if (this.OnLog == null)
        return;
      // ISSUE: reference to a compiler-generated field
      this.OnLog(msg);
    }

    public void LoadWeb(string address)
    {
      this._brouserComponent.Load(address);
          
    }
    private void _brouserComponent_LoadingStateChanged(object sender, LoadingStateChangedEventArgs e)
    {
      if (e.IsLoading)
        return;
      switch (this._curStatus)
      {
        case Status.Loading:
          this._curStatus = Status.Loggin;
          this.Log("Loging user..");
          this.LoginUser();
          isLogged = true;
          break;
        case Status.Loggin:
          this._curStatus = Status.Searching;
          this.Log("Going to AgentPerformance Report page....");
          this.GoToOrdersPage();
          break;
        case Status.Searching:
          this._curStatus = Status.GetAgentsInfo;
          this.Log("Getting Data....");
          if (!((ChromiumWebBrowser) sender).Address.Equals(this._newAgentsInfoSite))
            break;
          this.GetAgentPerformanceInfo();
          break;
        case Status.GetAgentsInfo:
          if (!((ChromiumWebBrowser) sender).Address.Equals(this._newAgentsInfoSite))
            break;
          this.GetAgentPerformanceInfo();
          break;
      }
    }

    public string GetColumnValue(List<List<string>> listInfo, List<string> agent, string columnName)
    {
      int index = listInfo[0].FindIndex((Predicate<string>) (a => a.Equals(columnName)));
      return index >= 0 ? agent[index] : string.Empty;
    }

    private List<AgentsInfo> ConvertToAgentList(List<List<string>> listInfo)
    {
      List<AgentsInfo> agentsInfoList = new List<AgentsInfo>();
      try
      {
        int num = 0;
        foreach (List<string> agent in listInfo)
        {
          if ((uint) num > 0U)
          {
            AgentsInfo agentsInfo = new AgentsInfo();
            int result = 0;
            string[] strArray = agent[0].Split('-');
            if (strArray.Length > 1)
              int.TryParse(strArray[1], out result);
            agentsInfo.USER = strArray[0].Trim();
            agentsInfo.UserID = new int?(result);
            agentsInfo.USERGROUP = this.GetColumnValue(listInfo, agent, "USER GROUP");
            agentsInfo.CALLS = this.GetColumnValue(listInfo, agent, "CALLS");
            agentsInfo.TIME = this.GetColumnValue(listInfo, agent, "TIME");
            agentsInfo.PAUSE = this.GetColumnValue(listInfo, agent, "PAUSE");
            agentsInfo.PAUSEAVG = this.GetColumnValue(listInfo, agent, "PAUSE AVG");
            agentsInfo.WAIT = this.GetColumnValue(listInfo, agent, "WAIT");
            agentsInfo.WAITAVG = this.GetColumnValue(listInfo, agent, "WAIT AVG");
            agentsInfo.TALK = this.GetColumnValue(listInfo, agent, "TALK");
            agentsInfo.TALKAVG = this.GetColumnValue(listInfo, agent, "TALK AVG");
            agentsInfo.DISPO = this.GetColumnValue(listInfo, agent, "DISPO");
            agentsInfo.DISPOAVG = this.GetColumnValue(listInfo, agent, "DISPO AVG");
            agentsInfo.DEAD = this.GetColumnValue(listInfo, agent, "DEAD");
            agentsInfo.DEADAVG = this.GetColumnValue(listInfo, agent, "DEAD AVG");
            agentsInfo.CUST = this.GetColumnValue(listInfo, agent, "CUST");
            agentsInfo.CUSTAVG = this.GetColumnValue(listInfo, agent, "CUST AVG");
            agentsInfo.X1 = this.GetColumnValue(listInfo, agent, "X1");
            agentsInfo.NORSPN = this.GetColumnValue(listInfo, agent, "NORSPN");
            agentsInfo.N = this.GetColumnValue(listInfo, agent, "N");
            agentsInfo.NRN = this.GetColumnValue(listInfo, agent, "NRN");
            agentsInfo.DNQ = this.GetColumnValue(listInfo, agent, "DNQ");
            agentsInfo.INSCH = this.GetColumnValue(listInfo, agent, "INSCH");
            agentsInfo.NOSCH = this.GetColumnValue(listInfo, agent, "NOSCH");
            agentsInfo.B = this.GetColumnValue(listInfo, agent, "B");
            agentsInfo.WRONG = this.GetColumnValue(listInfo, agent, "WRONG");
            agentsInfo.CBHOLD = this.GetColumnValue(listInfo, agent, "CBHOLD");
            agentsInfo.DNC = this.GetColumnValue(listInfo, agent, "DNC");
            agentsInfo.A = this.GetColumnValue(listInfo, agent, "A");
            agentsInfo.XHT = this.GetColumnValue(listInfo, agent, "XHT");
            agentsInfo.HS = this.GetColumnValue(listInfo, agent, "HS");
            agentsInfo.HU = this.GetColumnValue(listInfo, agent, "HU");
            agentsInfo.X2 = this.GetColumnValue(listInfo, agent, "X2");
            agentsInfo.CBL = this.GetColumnValue(listInfo, agent, "CBL");
            agentsInfo.NI = this.GetColumnValue(listInfo, agent, "NI");
            agentsInfo.X3 = this.GetColumnValue(listInfo, agent, "X3");
            agentsInfo.X4 = this.GetColumnValue(listInfo, agent, "X4");
            agentsInfo.JOB = this.GetColumnValue(listInfo, agent, "JOB");
            agentsInfo.XFER = this.GetColumnValue(listInfo, agent, "XFER");
            agentsInfo.NOT18 = this.GetColumnValue(listInfo, agent, "NOT18");
            agentsInfo.Date = DateTime.Now.ToShortDateString();
            agentsInfoList.Add(agentsInfo);
          }
          ++num;
        }
      }
      catch (Exception ex)
      {
      }
      return agentsInfoList;
    }

    private List<List<string>> ScrapeHtml(string website)
    {
      HtmlDocument htmlDocument = new HtmlDocument();
      htmlDocument.LoadHtml(website);
      return htmlDocument.DocumentNode.SelectSingleNode("//table").Descendants("tr").Skip<HtmlNode>(1).Where<HtmlNode>((Func<HtmlNode, bool>) (tr => tr.Elements("td").Count<HtmlNode>() > 1)).Select<HtmlNode, List<string>>((Func<HtmlNode, List<string>>) (tr => tr.Elements("td").Select<HtmlNode, string>((Func<HtmlNode, string>) (td => td.InnerText.Trim())).ToList<string>())).ToList<List<string>>();
    }

    private void GetAgentPerformanceInfo()
    {
      Thread.Sleep(4000);
      List<AgentsInfo> listAgent = new List<AgentsInfo>();
      this._brouserComponent.EvaluateScriptAsync("(function() {                              \r\n                               return document.documentElement.innerHTML;                                                    \r\n                             })();", new TimeSpan?()).ContinueWith((Action<Task<JavascriptResponse>>) (t =>
      {
        if (t.IsFaulted)
          return;
        listAgent = this.ConvertToAgentList(this.ScrapeHtml(t.Result.Result.ToString()));
        //Save agents info in the database
        this.SaveAgentsToDataBase(listAgent);
        this.SaveAgentTimestamp(listAgent);
          List<AgentsInfoExtended> listAgentsExtended = new List<AgentsInfoExtended>();
          var total = listAgent.Where(a => a.USER.Equals("Total")).FirstOrDefault();
          if (total != null)
          {
              listAgent.Remove(total);
          }

          foreach (var agent in listAgent)
          {
              var service = new AgentsPerformanceService();
              var agentExtended =  service.GetPerformanceData(DateTime.Now, agent.UserID);
              listAgentsExtended.Add(agentExtended);
          }
          foreach (var agent in listAgent)
          {
              SetEstatics(agent.USER, agent.UserID, listAgentsExtended);
          }
        //
        this._curStatus = Status.Completed;
        // ISSUE: reference to a compiler-generated field
        this.OnCompleted();
      }));
    }
        private string ConvertTimeToString(DateTime time)
        {
            var newTime = string.Format("{0}:{1}:{2}", time.Hour, time.Minute, time.Second);
            return newTime;
        }
        private double ConvertToTimeToDouble(string time)
        {
            try
            {
                if (string.IsNullOrEmpty(time))
                    time = "0:00:00";

                var splitHours = time.Split(':');

                int secondsTime = 0, minutesTime = 0, hoursTime = 0;

                bool talk = false, wait = false, dead = false;

                int.TryParse(splitHours[0], out hoursTime);
                int.TryParse(splitHours[1], out minutesTime);
                int.TryParse(splitHours[2], out secondsTime);

                if (secondsTime >= 60)
                {
                    minutesTime = minutesTime + 1;
                    secondsTime = 0;
                    talk = true;
                }
                if (minutesTime >= 60)
                {
                    hoursTime = hoursTime + 1;
                    minutesTime = 0;
                    talk = true;
                }

                TimeSpan Time;
                if (talk)
                {
                    var talkString = hoursTime.ToString() + ":" + minutesTime.ToString() + ":" + secondsTime.ToString();
                    Time = TimeSpan.Parse(talkString);
                }
                else
                {
                    Time = TimeSpan.Parse(time);
                }


                double hoursNew = Convert.ToDouble(Time.Hours) + Convert.ToDouble(Convert.ToDouble(Time.Minutes) / Convert.ToDouble(60));

                return Math.Round(hoursNew, 2);
            }
            catch (Exception ex)
            {

            }

            return 0;
        }

      private string GetActualRank(double lph, List<AgentsInfoExtended> listAgents, string timeUtilization)
      {
            var percent = timeUtilization.Replace("%", "");
            double percentConverted = 0.0;
            Double.TryParse(percent, out percentConverted);

            var ListRanks = new List<Rank>();
            var rank1 = new Rank
            {
                LPH = 999.9,
                UserId = 999,
                Username = "AgentRank1"
            };
            var rank2 = new Rank
            {
                LPH = 999.8,
                UserId = 998,
                Username = "AgentRank2"
            };
            ListRanks.Add(rank1);
            ListRanks.Add(rank2);

            if (lph <= 1.7)
              {
                  return "Not Ranked";
              }

            foreach (var agent in listAgents)
              {
                  var service = new AgentsPerformanceService();
                  Rank rank = new Rank();
                  rank.LPH = agent.LPH;
                  rank.UserId = agent.Id;
                  rank.Username = agent.USER;
                //Time Utilization
                var waitTime = TimeSpan.Parse(agent.WAIT);
                var talkTime = TimeSpan.Parse(agent.TALK);
                var deadTime = TimeSpan.Parse(agent.DEAD);
                var workingTime = (waitTime + talkTime + deadTime);
                var breakTime = TimeSpan.Parse("00:45:00");
                var totalBreakTime = new TimeSpan();
                var timeNow = ConvertTimeToString(DateTime.Now);
                var userTimeStamp = service.GetTodayTimeStamp(agent.USER, DateTime.Now);
                var agentsLoginTime = ConvertTimeToString(userTimeStamp.Value);

                var potentialWorkingHour = ConvertToTimeToDouble(timeNow) - ConvertToTimeToDouble(agentsLoginTime);
                //AgentPerformanceInfo agentPerformance = new AgentPerformanceInfo();
                //Calculate Break Time
                if (workingTime.Hours < 4)
                {
                    var substract = TimeSpan.Parse("00:35:00");
                    totalBreakTime = breakTime - substract;
                }
                else if (workingTime.Hours >= 4 && workingTime.Hours < 6)
                {
                    var substract = TimeSpan.Parse("00:20:00");
                    totalBreakTime = breakTime - substract;
                }
                else
                {
                    totalBreakTime = breakTime;
                }
                var timeUti = Math.Round(
                ((ConvertToTimeToDouble(workingTime.ToString()) +
                  ConvertToTimeToDouble(totalBreakTime.ToString())) / potentialWorkingHour), 2);

                rank.TimeUtilization = timeUti* 100;
                //
                ListRanks.Add(rank);
              }

              var newList = ListRanks.OrderByDescending(a => a.LPH).ToList();
              for (int i = 0; i < newList.Count; i++)
              {
                  if (newList[i].LPH.Equals(lph))
                  {
                      if (i == 2 && percentConverted >=90 )
                      {
                         return "1";
                      }
                      if (i == 3 && newList[2].TimeUtilization < 90 && newList[3].TimeUtilization >=90)
                      {
                         return "1";
                      }
                      if (i == 3 && percentConverted >= 90)
                      {
                         return "2";
                      }

                      return (i + 1).ToString();
                  }
              }
          //    scope.Complete();
          //}
          return "Not Ranked";
      }
        private void SetEstatics(string agentName, int? userId, List<AgentsInfoExtended> listAgents)
        {
            //var dialerId = 0;
            //int.TryParse(userId, out dialerId);
            try
            {
                    var service = new AgentsPerformanceService();
                    var userData = listAgents.FirstOrDefault(a => a.Id.Equals(userId));
                    var workingHours = service.GetHours(userData);

                    var userTimeStamp = service.GetTodayTimeStamp(agentName, DateTime.Now);

                    //Return if null
                    if (userData == null) return;

                    var waitTime = TimeSpan.Parse(userData.WAIT);
                    var talkTime = TimeSpan.Parse(userData.TALK);
                    var deadTime = TimeSpan.Parse(userData.DEAD);
                    var workingTime = (waitTime + talkTime + deadTime);
                    var breakTime = TimeSpan.Parse("00:45:00");
                    var totalBreakTime = new TimeSpan();
                    var timeNow = ConvertTimeToString(DateTime.Now);
                    var agentsLoginTime = ConvertTimeToString(userTimeStamp.Value);

                    var potentialWorkingHour = ConvertToTimeToDouble(timeNow) - ConvertToTimeToDouble(agentsLoginTime);
                    AgentPerformanceInfo agentPerformance = new AgentPerformanceInfo();

                    //Working time
                    agentPerformance.EstWorkingTime = Math.Round(workingHours, 2).ToString();

                    //Total LPH
                    agentPerformance.TotalLPH = Math.Round(userData.LPH, 3).ToString();

                    //Calculate Break Time
                    if (workingTime.Hours < 4)
                    {
                        var substract = TimeSpan.Parse("00:35:00");
                        totalBreakTime = breakTime - substract;
                    }
                    else if (workingTime.Hours >= 4 && workingTime.Hours < 6)
                    {
                        var substract = TimeSpan.Parse("00:20:00");
                        totalBreakTime = breakTime - substract;
                    }
                    else
                    {
                        totalBreakTime = breakTime;
                    }
                    var timeUti = Math.Round(
                    ((ConvertToTimeToDouble(workingTime.ToString()) +
                      ConvertToTimeToDouble(totalBreakTime.ToString()))/potentialWorkingHour), 2);
                    if (timeUti > 1)
                    {
                        timeUti = 1.0;
                    }
                    agentPerformance.TimeUtilization = (timeUti*100).ToString() + "%";


                    agentPerformance.Rank = timeUti > 0.9
                        ? Math.Round(userData.LPH, 2).ToString()
                        : (userData.LPH*timeUti).ToString();

                    agentPerformance.TotalNoXFER = userData.XFER;
                    //Total Leads
                    var x1 = 0;
                    int x2 = 0;
                    int x3 = 0;
                    int x4 = 0;
                    int xht = 0;
                    int lph = 0;
                    int.TryParse(userData.X1, out x1);
                    int.TryParse(userData.X2, out x2);
                    int.TryParse(userData.X3, out x3);
                    int.TryParse(userData.X4, out x4);
                    int.TryParse(userData.XHT, out xht);
                    //int.TryParse(userData.LPH, out lph);

                    agentPerformance.TotalLeads = ((x1) + (x2*2) + (x3*3) + (x4*4) + (xht*3)).ToString();
                    agentPerformance.TotalProjectedLead = ((8*userData.LPH)*.9).ToString();
                    var scrapingEntities = new SchoolEntities();
                    var todayDate = DateTime.Now.ToShortDateString();
                    var agent = scrapingEntities.AgentPerformanceInfoes
                        .FirstOrDefault(a => a.Date.Equals(todayDate) && a.UserID.Equals(userId.Value));

                    if (agent != null)
                    {
                        agent.TotalLeads = agentPerformance.TotalLeads;
                        agent.Date = DateTime.Now.ToShortDateString();
                        agent.EstWorkingTime = agentPerformance.EstWorkingTime;
                        agent.Rank = GetActualRank(userData.LPH, listAgents, agentPerformance.TimeUtilization); //agentPerformance.Rank;
                        agent.TimeUtilization = agentPerformance.TimeUtilization;
                        agent.TotalLPH = agentPerformance.TotalLPH;
                        agent.TotalNoXFER = agentPerformance.TotalNoXFER;
                        agent.TotalProjectedLead = agentPerformance.TotalProjectedLead;
                        agent.UserID = userId.Value;
                        agent.UserName = agentName;

                    }
                    else
                    {
                        agent = new Appraisal.AgentPerformanceInfo();
                        agent.TotalLeads = agentPerformance.TotalLeads;
                        agent.Date = DateTime.Now.ToShortDateString();
                        agent.EstWorkingTime = agentPerformance.EstWorkingTime;
                        agent.Rank = GetActualRank(userData.LPH, listAgents, agentPerformance.TimeUtilization);
                        agent.TimeUtilization = agentPerformance.TimeUtilization;
                        agent.TotalLPH = agentPerformance.TotalLPH;
                        agent.TotalNoXFER = agentPerformance.TotalNoXFER;
                        agent.TotalProjectedLead = agentPerformance.TotalProjectedLead;
                        agent.UserID = userId.Value;
                        agent.UserName = agentName;
                    }

                    scrapingEntities.AgentPerformanceInfoes.AddOrUpdate(agent);
                    scrapingEntities.SaveChanges();

                  //  scope.Complete();
                //}

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }

        }
        public void SaveAgentTimestamp(List<AgentsInfo> agentsInfo )
      {
          try
          {
               var now = DateTime.Now;
              if (now.Hour <= 7 || now.Hour >= 22) return;

              var scrapingEntities = new SchoolEntities();
              var todayDate = DateTime.Now;
              foreach (var agent in agentsInfo)
              {
                  var exits =
                      scrapingEntities.AgentTimestamps.Any(
                          a =>
                              EntityFunctions.TruncateTime(a.Datetime) == EntityFunctions.TruncateTime(todayDate) &&
                              a.UserName.Equals(agent.USER));
                  if (!exits)
                  {
                      var agentTime = new AgentTimestamp
                      {
                          Datetime = DateTime.Now,
                          UserID = agent.UserID ?? 0,
                          UserName = agent.USER
                      };
                      scrapingEntities.AgentTimestamps.Add(agentTime);
                  }
                  else
                  {

                  }
              }
              scrapingEntities.SaveChanges();
              //    scope.Complete();
                //}
              
          }
          catch (Exception ex)
          {
              //MessageBox.Show(ex.StackTrace);
          }
      }
    private void SaveAgentsToDataBase(List<AgentsInfo> listAgent)
    {
      try
      {
        DateTime now = DateTime.Now;
        int num;
        if (now.DayOfWeek != DayOfWeek.Saturday)
        {
          now = DateTime.Now;
          if (now.DayOfWeek != DayOfWeek.Sunday)
          {
            now = DateTime.Now;
            if (now.Hour >= 7)
            {
              now = DateTime.Now;
              num = now.Hour > 22 ? 1 : 0;
              goto label_6;
            }
          }
        }
        num = 1;
    label_6:
        if (num != 0)
          return;
        SchoolEntities scrapingEntities = new SchoolEntities();
        now = DateTime.Now;
        string todayDate = now.ToShortDateString();
        IQueryable<AgentsInfo> source = scrapingEntities.AgentsInfoes.Where<AgentsInfo>((Expression<Func<AgentsInfo, bool>>) (a => a.Date.Equals(todayDate)));
        using (TransactionScope transactionScope = new TransactionScope())
        {
          if (listAgent.Count <= 0)
            return;
          if (source.Any<AgentsInfo>())
          {
            foreach (AgentsInfo entity in (IEnumerable<AgentsInfo>) source)
              scrapingEntities.AgentsInfoes.Remove(entity);
          }
          foreach (AgentsInfo entity in listAgent)
            scrapingEntities.AgentsInfoes.Add(entity);
          scrapingEntities.SaveChanges();
          transactionScope.Complete();
        }
      }
      catch (Exception ex)
      {
      }
    }

    private void GoToOrdersPage()
    {
      this._brouserComponent.Load(this._newAgentsInfoSite);
    }

    public void Dispose()
    {
    }

    public delegate void LogMsg(string msg);

    public delegate void Completed();
  }
}
