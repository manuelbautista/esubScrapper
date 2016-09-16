using System;
using System.Collections.Generic;
using System.Data.Objects;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Dax.Scrapping.Appraisal.Core
{
    public  class AgentsPerformanceService
    {
        public List<string> GetAgents()
        {
            var school = new SchoolEntities();
            var agents = school.View_Agents.Select(a => a.USER);

            return agents.ToList();
        }

        public AgentsInfoExtended ConvertToAgentExdented(AgentsInfo agent)
        {
            var agentExtended = new AgentsInfoExtended();
            agentExtended.Id = agent.UserID.Value;
            agentExtended.USER = agent.USER;
            agentExtended.USERGROUP = agent.USERGROUP;
            agentExtended.CALLS = agent.CALLS;
            agentExtended.TIME = agent.TIME;
            agentExtended.PAUSE = agent.PAUSE;
            agentExtended.PAUSEAVG = agent.PAUSEAVG;
            agentExtended.WAIT = agent.WAIT;
            agentExtended.WAITAVG = agent.WAITAVG;
            agentExtended.TALK = agent.TALK;
            agentExtended.TALKAVG = agent.TALKAVG;
            agentExtended.DISPO = agent.DISPO;
            agentExtended.DISPOAVG = agent.DISPOAVG;
            agentExtended.DEAD = agent.DEAD;
            agentExtended.DEADAVG = agent.DEADAVG;
            agentExtended.CUST = agent.CUST;
            agentExtended.CUSTAVG = agent.CUSTAVG;
            agentExtended.X1 = agent.X1;
            agentExtended.NORSPN = agent.NORSPN;
            agentExtended.N = agent.N;
            agentExtended.NRN = agent.NRN;
            agentExtended.DNQ = agent.DNQ;
            agentExtended.INSCH = agent.INSCH;
            agentExtended.NOSCH = agent.NOSCH;
            agentExtended.B = agent.B;
            agentExtended.WRONG = agent.WRONG;
            agentExtended.CBHOLD = agent.CBHOLD;
            agentExtended.DNC = agent.DNC;
            agentExtended.A = agent.A;
            agentExtended.XHT = agent.XHT;
            agentExtended.HS = agent.HS;
            agentExtended.HU = agent.HU;
            agentExtended.X2 = agent.X2;
            agentExtended.CBL = agent.CBL;
            agentExtended.NI = agent.NI;
            agentExtended.X3 = agent.X3;
            agentExtended.X4 = agent.X4;
            agentExtended.JOB = agent.JOB;
            agentExtended.XFER = agent.XFER;
            agentExtended.NOT18 = agent.NOT18;
            agentExtended.Date = agent.Date;
            agentExtended.TotalUnique = 0;
            agentExtended.TotalSchools = 0;
            agentExtended.TotalEDS = 0;
            agentExtended.PercentofTransfers = 0.0;
            agentExtended.MatchRate = 0.0;
            agentExtended.Hours = 0.0;
            agentExtended.UniqueConversionRate = 0.0;
            agentExtended.TotalConversionRate = 0.0;
            agentExtended.LPH = 0.0;
            agentExtended.UPH = 0.0;
            agentExtended.CPH = 0.0;
            return agentExtended;
        }

        public int GetTotalUnique(AgentsInfoExtended agent)
        {
            int x1 = 0, x2 = 0, x3 = 0, x4 = 0, xht = 0;
            var totalUnique = 0;
            int.TryParse(agent.X1, out x1);
            int.TryParse(agent.X2, out x2);
            int.TryParse(agent.X3, out x3);
            int.TryParse(agent.X4, out x4);
            int.TryParse(agent.XHT, out xht);

            totalUnique = x1 + x2 + x3 + x4 + xht;

            return totalUnique;
        }

        public string GetCallsPercent(DateTime date, string calls)
        {
            var school = new SchoolEntities();
            var agentsInfos = school.AgentsInfoes.Select(a => a).ToList();
            var initDate = date.ToShortDateString();
            var initDate2 = Convert.ToDateTime(initDate);
            var endDate = date.ToShortDateString();
            var endDate2 = Convert.ToDateTime(endDate).AddHours(23);

            var agents = agentsInfos.Where(a => Convert.ToDateTime(a.Date) >= initDate2 && Convert.ToDateTime(a.Date) <= endDate2).ToList();
            int agentsQty = agents.Count;
            var callsOrdered = agents.OrderByDescending(a => int.Parse(a.CALLS));
            int pos = 0;
            foreach (var agent in callsOrdered)
            {
                if (agent.CALLS == calls)
                {
                    break;
                }
                else
                {
                    pos++;
                }

            }
            return string.Format("{0} / {1}", pos, agentsQty);
        }

        public string GetLPHPercent(DateTime date, double LPH)
        {
            var agents = GetPerformanceByDatesAndUser(date, date, "----All----");
            int agentsQty = agents.Count();
            var callsOrdered = agents.OrderByDescending(a => a.LPH);
            int pos = 0;
            foreach (var agent in callsOrdered)
            {
                if (agent.LPH <= LPH)
                {
                    break;
                }
                else
                {
                    pos++;
                }

            }
            return string.Format("{0} / {1}", pos, agentsQty);
        }

        public string GetTotalEDSPercent(DateTime date, double totalEDS)
        {
            var agents = GetPerformanceByDatesAndUser(date, date, "----All----");
            int agentsQty = agents.Count();
            var callsOrdered = agents.OrderByDescending(a => a.TotalEDS);
            int pos = 0;
            foreach (var agent in callsOrdered)
            {
                if (agent.TotalEDS <= totalEDS)
                {
                    break;
                }
                else
                {
                    pos++;
                }

            }
            return string.Format("{0} / {1}", pos, agentsQty);
        }

        public string GetTotalSchoolsPercent(DateTime date, double totalSchools)
        {
            var agents = GetPerformanceByDatesAndUser(date, date, "----All----");
            int agentsQty = agents.Count();
            var callsOrdered = agents.OrderByDescending(a => a.TotalSchools);
            int pos = 0;
            foreach (var agent in callsOrdered)
            {
                if (agent.TotalSchools <= totalSchools)
                {
                    break;
                }
                else
                {
                    pos++;
                }

            }
            return string.Format("{0} / {1}", pos, agentsQty);
        }

        public string GetCPHPercent(DateTime date, double CPH)
        {
            var agents = GetPerformanceByDatesAndUser(date, date, "----All----");
            int agentsQty = agents.Count();
            var callsOrdered = agents.OrderByDescending(a => a.CPH);
            int pos = 0;
            foreach (var agent in callsOrdered)
            {
                if (agent.CPH <= CPH)
                {
                    break;
                }
                else
                {
                    pos++;
                }

            }
            return string.Format("{0} / {1}", pos, agentsQty);
        }
        public int GetTotalSchools(AgentsInfoExtended agent)
        {
            int x1 = 0, x2 = 0, x3 = 0, x4 = 0, xht = 0;
            var totalSchools = 0;
            int.TryParse(agent.X1, out x1);
            int.TryParse(agent.X2, out x2);
            int.TryParse(agent.X3, out x3);
            int.TryParse(agent.X4, out x4);
            int.TryParse(agent.XHT, out xht);

            x2 = x2 * 2;
            x3 = x3 * 3;
            x4 = x4 * 4;
            xht = xht * 3;

            totalSchools = x1 + x2 + x3 + x4 + xht;

            return totalSchools;
        }

        public double GetHours(AgentsInfoExtended agent)
        {
            if (agent.Id > 0)
            {
                try
                {
                    if (string.IsNullOrEmpty(agent.TALK))
                        agent.TALK = "0:00";
                    if (string.IsNullOrEmpty(agent.WAIT))
                        agent.WAIT = "0:00";
                    if (string.IsNullOrEmpty(agent.DEAD))
                        agent.DEAD = "0:00";

                    var splitHoursTalk = agent.TALK.Split(':');
                    var splitHoursWait = agent.WAIT.Split(':');
                    var splitHoursDead = agent.DEAD.Split(':');
                    int secondsTalk = 0, minutesTalk = 0, hoursTalk = 0;
                    int secondsWait = 0, minutesWait = 0, hoursWait = 0;
                    int secondsDead = 0, minutesDead = 0, hoursDead = 0;
                    bool talk = false, wait = false, dead = false;

                    int.TryParse(splitHoursTalk[0], out hoursTalk);
                    int.TryParse(splitHoursTalk[1], out minutesTalk);
                    int.TryParse(splitHoursTalk[2], out secondsTalk);

                    int.TryParse(splitHoursWait[0], out hoursWait);
                    int.TryParse(splitHoursWait[1], out minutesWait);
                    int.TryParse(splitHoursWait[2], out secondsWait);

                    int.TryParse(splitHoursDead[0], out hoursDead);
                    int.TryParse(splitHoursDead[1], out minutesDead);
                    int.TryParse(splitHoursDead[2], out secondsDead);

                    if (secondsTalk >= 60)
                    {
                        minutesTalk = minutesTalk + 1;
                        secondsTalk = 0;
                        talk = true;
                    }
                    if (minutesTalk >= 60)
                    {
                        hoursTalk = hoursTalk + 1;
                        minutesTalk = 0;
                        talk = true;
                    }

                    if (secondsWait >= 60)
                    {
                        minutesWait = minutesWait + 1;
                        secondsWait = 0;
                        wait = true;
                    }
                    if (minutesWait >= 60)
                    {
                        hoursWait = hoursWait + 1;
                        minutesWait = 0;
                        wait = true;
                    }

                    if (secondsDead >= 60)
                    {
                        minutesDead = minutesDead + 1;
                        secondsDead = 0;
                        dead = true;
                    }
                    if (minutesDead >= 60)
                    {
                        hoursDead = hoursDead + 1;
                        minutesDead = 0;
                        dead = true;
                    }

                    TimeSpan Talk;
                    TimeSpan Wait;
                    TimeSpan Dead;
                    if (talk)
                    {
                        var talkString = hoursTalk.ToString() + ":" + minutesTalk.ToString() + ":" + secondsTalk.ToString();
                        Talk = TimeSpan.Parse(talkString);
                    }
                    else
                    {
                        Talk = TimeSpan.Parse(agent.TALK);
                    }

                    if (wait)
                    {
                        var waitString = hoursWait.ToString() + ":" + minutesWait.ToString() + ":" + secondsWait.ToString();
                        Wait = TimeSpan.Parse(waitString);
                    }
                    else
                    {
                        Wait = TimeSpan.Parse(agent.WAIT);
                    }

                    if (dead)
                    {
                        var deadString = hoursDead.ToString() + ":" + minutesDead.ToString() + ":" + secondsDead.ToString();
                        Dead = TimeSpan.Parse(deadString);
                    }
                    else
                    {
                        Dead = TimeSpan.Parse(agent.DEAD);
                    }


                    var hours = (Talk + Wait) - Dead;

                    double hoursNew = Convert.ToDouble(hours.Hours) + Convert.ToDouble(Convert.ToDouble(hours.Minutes) / Convert.ToDouble(60));

                    return hoursNew;
                }
                catch (Exception ex)
                {

                }
            }
            return 0;
        }

        public List<View_Agents> GetUserIdByUserNames(List<string> agents)
        {
            var school = new SchoolEntities();
            var Users = school.View_Agents.Where(a => agents.Contains(a.USER.ToLower()));

            return Users.ToList();
        }

        public DateTime? GetTodayTimeStamp(string agentName, DateTime date)
        {
            var school = new SchoolEntities();

            var timestamp = school.AgentTimestamps.FirstOrDefault(a => a.UserName.Equals(agentName) && EntityFunctions.TruncateTime(a.Datetime) == EntityFunctions.TruncateTime(date));
            if (timestamp != null)
            {
                return timestamp.Datetime;
            }
            return DateTime.Now;
        }

        public AgentsInfoExtended GetPerformanceData(DateTime init, int? userId)
        {
            try
            {
                var school = new SchoolEntities();
                var todayDate = DateTime.Now.ToShortDateString();
                var id = userId.Value;

                List<AgentsInfoExtended> listAgentsExtended = new List<AgentsInfoExtended>();
                var agent = new AgentsInfo();

                var agentsInfos = school.AgentsInfoes.Select(a => a).Where(a => a.Date.Equals(todayDate) && a.UserID.HasValue && a.UserID.Value.Equals(id)).OrderByDescending(a => a.USER).FirstOrDefault();
                agent = agentsInfos;
                if (agent != null)
                {
                    TimeZoneInfo easternStandardTimeTimeZoneInfo =
                        TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                    DateTime utcNow = DateTime.UtcNow;
                    DateTime easternStandardTime = TimeZoneInfo.ConvertTimeFromUtc(utcNow,
                        easternStandardTimeTimeZoneInfo);

                    double xfer = 0;
                    double.TryParse(agent.XFER, out xfer);
                    int calls = 0;
                    int.TryParse(agent.CALLS, out calls);
                    var agentExtended = ConvertToAgentExdented(agent);
                    //var val1 = (agentExtended.TotalUnique + xfer);
                    var hours = GetHours(agentExtended);

                    agentExtended.TotalUnique = GetTotalUnique(agentExtended);
                    agentExtended.TotalSchools = GetTotalSchools(agentExtended);
                    agentExtended.TotalEDS = xfer;
                    if ((agentExtended.TotalUnique + xfer) <= 0)
                    {
                        agentExtended.PercentofTransfers = 0.0;
                    }
                    else
                    {
                        agentExtended.PercentofTransfers =
                        (Math.Round(Convert.ToDouble(xfer / (agentExtended.TotalUnique + xfer)), 4) *
                         Convert.ToDouble(100.00));
                    }

                    agentExtended.MatchRate = agentExtended.TotalUnique > 0
                        ? Math.Round(
                            Convert.ToDouble(Convert.ToDouble(agentExtended.TotalSchools) /
                                             Convert.ToDouble(agentExtended.TotalUnique)), 2)
                        : 0.0;
                    agentExtended.Hours = Math.Round(Convert.ToDouble(GetHours(agentExtended)), 2);
                    agentExtended.UniqueConversionRate = calls > 0
                        ? (Math.Round(
                               Convert.ToDouble(Convert.ToDouble(agentExtended.TotalUnique) / Convert.ToDouble(calls)), 4) *
                           Convert.ToDouble(100.00))
                        : 0.0;
                    agentExtended.TotalConversionRate = calls > 0
                        ? (Math.Round(
                               Convert.ToDouble(Convert.ToDouble(agentExtended.TotalSchools) / Convert.ToDouble(calls)), 4) *
                           Convert.ToDouble(100.00))
                        : 0.0;
                    agentExtended.LPH = hours > 0
                        ? Math.Round(
                            Convert.ToDouble(Convert.ToDouble(agentExtended.TotalSchools) / Convert.ToDouble(hours)), 3)
                        : 0.0;
                    agentExtended.UPH = hours > 0
                        ? Math.Round(
                            Convert.ToDouble((Convert.ToDouble(agentExtended.TotalUnique)) /
                                             Convert.ToDouble(GetHours(agentExtended))), 2)
                        : 0.0;
                    agentExtended.CPH = hours > 0
                        ? Math.Round(
                            Convert.ToDouble(Convert.ToDouble(calls) / Convert.ToDouble(GetHours(agentExtended))), 2)
                        : 0.0;

                    if (!agent.USER.ToLower().Equals("total"))
                    {
                        DateTime date = DateTime.Parse(agent.Date);
                        TimeSpan currentTime = date == DateTime.Now.Date
                            ? easternStandardTime.TimeOfDay
                            : currentTime = GetAgentLastLogout(agent.USER, date);
                        TimeSpan firstLogin = GetAgentFirstLogin(agent.USER, date);
                        TimeSpan mpt = GetAgentMPT(agent.USER, date);
                        TimeSpan time = TimeSpan.Parse(agent.TIME);
                        TimeSpan pause = TimeSpan.Parse(agent.PAUSE);
                        TimeSpan logoutTime = currentTime.Subtract(firstLogin).Subtract(time).Subtract(mpt);
                        if (logoutTime < TimeSpan.Zero)
                        {
                            logoutTime = TimeSpan.Parse("00:00:00");
                        }

                        TimeSpan realPause = pause.Add(logoutTime);
                        agentExtended.MPT = mpt.ToString(@"hh\:mm\:ss");
                        agentExtended.LogoutTime = logoutTime.ToString(@"hh\:mm\:ss");
                        agentExtended.RealPause = realPause.ToString(@"hh\:mm\:ss");
                    }

                    return agentExtended;
                }
            }
            catch (Exception ex)
            {

                //MessageBox.Show(ex.StackTrace);
            }

            return new AgentsInfoExtended();
        }
        public IEnumerable<AgentsInfoExtended> GetPerformanceByDatesAndUser(DateTime init, DateTime end, string user, string agentGroup = "")
        {
            try
            {
                var school = new SchoolEntities();
                var initDate = init.ToShortDateString();
                var initDate2 = Convert.ToDateTime(initDate);
                var endDate = end.ToShortDateString();
                var endDate2 = Convert.ToDateTime(endDate).AddHours(23);
                List<AgentsInfoExtended> listAgentsExtended = new List<AgentsInfoExtended>();
                var agents = new List<AgentsInfo>();

                var agentsInfos = school.AgentsInfoes.Select(a => a).ToList();
                agents = user.Equals("----All----") ? agentsInfos.Where(a => Convert.ToDateTime(a.Date) >= initDate2 && Convert.ToDateTime(a.Date) <= endDate2).ToList() : agentsInfos.Where(a => Convert.ToDateTime(a.Date) >= initDate2 && Convert.ToDateTime(a.Date) <= endDate2 && a.UserID.Value > 0 && a.USER.ToLower().Equals(user.ToLower())).OrderBy(a => a.USER).ToList();
                if (!agentGroup.Equals("-------All-------"))
                {
                    agents = agents.Where(a => a.USERGROUP.ToLower().Equals(agentGroup.ToLower())).ToList();
                }

                TimeZoneInfo easternStandardTimeTimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Eastern Standard Time");
                DateTime utcNow = DateTime.UtcNow;
                DateTime easternStandardTime = TimeZoneInfo.ConvertTimeFromUtc(utcNow, easternStandardTimeTimeZoneInfo);
                foreach (var agent in agents)
                {
                    double xfer = 0;
                    double.TryParse(agent.XFER, out xfer);
                    int calls = 0;
                    int.TryParse(agent.CALLS, out calls);
                    var agentExtended = ConvertToAgentExdented(agent);
                    //var val1 = (agentExtended.TotalUnique + xfer);
                    var hours = GetHours(agentExtended);

                    agentExtended.TotalUnique = GetTotalUnique(agentExtended);
                    agentExtended.TotalSchools = GetTotalSchools(agentExtended);
                    agentExtended.TotalEDS = xfer;
                    if ((agentExtended.TotalUnique + xfer) <= 0)
                    {
                        agentExtended.PercentofTransfers = 0.0;
                    }
                    else
                    {
                        agentExtended.PercentofTransfers =
                            (Math.Round(Convert.ToDouble(xfer / (agentExtended.TotalUnique + xfer)), 4) * Convert.ToDouble(100.00));
                    }

                    agentExtended.MatchRate = agentExtended.TotalUnique > 0 ? Math.Round(Convert.ToDouble(Convert.ToDouble(agentExtended.TotalSchools) / Convert.ToDouble(agentExtended.TotalUnique)), 2) : 0.0;
                    agentExtended.Hours = Math.Round(Convert.ToDouble(GetHours(agentExtended)), 2);
                    agentExtended.UniqueConversionRate = calls > 0 ? (Math.Round(Convert.ToDouble(Convert.ToDouble(agentExtended.TotalUnique) / Convert.ToDouble(calls)), 4) * Convert.ToDouble(100.00)) : 0.0;
                    agentExtended.TotalConversionRate = calls > 0 ? (Math.Round(Convert.ToDouble(Convert.ToDouble(agentExtended.TotalSchools) / Convert.ToDouble(calls)), 4) * Convert.ToDouble(100.00)) : 0.0;
                    agentExtended.LPH = hours > 0 ? Math.Round(Convert.ToDouble(Convert.ToDouble(agentExtended.TotalSchools) / Convert.ToDouble(hours)), 2) : 0.0;
                    agentExtended.UPH = hours > 0 ? Math.Round(Convert.ToDouble((Convert.ToDouble(agentExtended.TotalUnique)) / Convert.ToDouble(GetHours(agentExtended))), 2) : 0.0;
                    agentExtended.CPH = hours > 0 ? Math.Round(Convert.ToDouble(Convert.ToDouble(calls) / Convert.ToDouble(GetHours(agentExtended))), 2) : 0.0;

                    if (!agent.USER.ToLower().Equals("total"))
                    {
                        DateTime date = DateTime.Parse(agent.Date);
                        TimeSpan currentTime = date == DateTime.Now.Date
                            ? easternStandardTime.TimeOfDay
                            : currentTime = GetAgentLastLogout(agent.USER, date);
                        TimeSpan firstLogin = GetAgentFirstLogin(agent.USER, date);
                        TimeSpan mpt = GetAgentMPT(agent.USER, date);
                        TimeSpan time = TimeSpan.Parse(agent.TIME);
                        TimeSpan pause = TimeSpan.Parse(agent.PAUSE);
                        TimeSpan logoutTime = currentTime.Subtract(firstLogin).Subtract(time).Subtract(mpt);
                        if (logoutTime < TimeSpan.Zero)
                        {
                            logoutTime = TimeSpan.Parse("00:00:00");
                        }

                        TimeSpan realPause = pause.Add(logoutTime);
                        agentExtended.MPT = mpt.ToString(@"hh\:mm\:ss");
                        agentExtended.LogoutTime = logoutTime.ToString(@"hh\:mm\:ss");
                        agentExtended.RealPause = realPause.ToString(@"hh\:mm\:ss");
                    }
                    listAgentsExtended.Add(agentExtended);
                }

                return listAgentsExtended;
            }
            catch (Exception ex)
            {


            }

            return new List<AgentsInfoExtended>();
        }

        private TimeSpan GetAgentMPT(string userName, DateTime date)
        {
            TimeSpan agentMPT;
            SchoolEntities schoolEntities = new SchoolEntities();
            var agentMPTQuery = from mpt in schoolEntities.AgentsMPTs
                                where (mpt.UserName == userName) && (mpt.Date == date)
                                select mpt.MPT;
            var agentMPTString = agentMPTQuery.SingleOrDefault();
            if (agentMPTString != null)
            {
                agentMPT = TimeSpan.Parse(agentMPTString);
            }
            else
            {
                agentMPT = TimeSpan.Parse("00:00:00");
            }

            return agentMPT;
        }

        private TimeSpan GetAgentFirstLogin(string userName, DateTime date)
        {
            SchoolEntities schoolEntities = new SchoolEntities();
            DateTime? firstLoginDateTime;
            if (date.Date == DateTime.Now.Date)
            {
                firstLoginDateTime = (from fl in schoolEntities.users
                                      where (fl.username == userName)
                                      select fl.FirstLogin).FirstOrDefault();
            }
            else
            {
                firstLoginDateTime = (from fl in schoolEntities.UsersHistories
                                      where (fl.UserName == userName) && (fl.Date == date)
                                      select fl.FirstLogin).FirstOrDefault();
            }

            TimeSpan firstLogin = TimeSpan.Parse("00:00:00");
            if (firstLoginDateTime.HasValue)
                if (firstLoginDateTime.Value != DateTime.MinValue)
                {
                    firstLogin = firstLoginDateTime.Value.TimeOfDay;
                }

            return firstLogin;
        }

        private TimeSpan GetAgentLastLogout(string userName, DateTime date)
        {
            SchoolEntities schoolEntities = new SchoolEntities();
            DateTime? lastLogoutDateTime;
            if (date.Date == DateTime.Now.Date)
            {
                lastLogoutDateTime = (from fl in schoolEntities.users
                                      where (fl.username == userName)
                                      select fl.LastLogout).FirstOrDefault();
            }
            else
            {
                lastLogoutDateTime = (from fl in schoolEntities.UsersHistories
                                      where (fl.UserName == userName) && (fl.Date == date)
                                      select fl.LastLogout).FirstOrDefault();
            }

            TimeSpan lastLogout = TimeSpan.Parse("00:00:00");
            if (lastLogoutDateTime.HasValue)
            {
                lastLogout = lastLogoutDateTime.Value.TimeOfDay;
            }

            return lastLogout;
        }
    }
}
