// Decompiled with JetBrains decompiler
// Type: Dax.Scrapping.Appraisal.MainForm
// Assembly: Dax.Scrapping.Appraisal, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 636504A0-34C1-4208-9BC8-AA1181CC609A
// Assembly location: C:\Users\manue\Desktop\EsubScraper\Dax.Scrapping.Appraisal.exe

using CefSharp;
using CefSharp.WinForms;
using Dax.Scrapping.Appraisal;
using Dax.Scrapping.Appraisal.Core;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data.Entity.Migrations;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Threading;
using System.Timers;
using System.Windows.Forms;
using ActiveUp.Net.Mail;
using Newtonsoft.Json;

namespace Dax.Scrapping.Appraisal
{
    public class PositionClass
    {
        public string X { get; set; }
        public string Y { get; set; }
    }
  public class MainForm : Form
  {
    private AppraisalScrapper _scraper = (AppraisalScrapper) null;
    private bool _AutoStart = true;
    private System.Timers.Timer aTimer = new System.Timers.Timer();
    private RegistryKey rkApp = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
    private IContainer components = (IContainer) null;
    private TabControl ConsultTabControl;
    private TabPage tabPage1;
    private TabPage tabPage2;
    private GroupBox groupBox1;
    private TextBox txtUserId;
    private Label label2;
    private Label label1;
    private TextBox txtUserPassword;
    private GroupBox groupBox3;
    private DataGridView dgvData;
    private BindingSource bindingSource1;
    private Button btnSearch;
    private TableLayoutPanel tableLayoutPanel1;
    private ProgressBar progressBar1;
    private Panel panel1;
    private RichTextBox txtLogs;
    private Panel panelBrowser;
    private WebBrowser webBrowser1;
    private Button buttonDevTool;
        private TabPage tabPage3;
        private Panel panelReportsDownload;
        private Panel panel2;
        private ListBox listBoxReportsDownload;
        private NotifyIcon notifyIcon1;
        private TabPage tabPage4;
        private Panel panelReport2;
        public bool isLogin = true;
      private ChromiumWebBrowser browserReport2Dev;
        private TabPage tabPage5;
        private Panel panelReport3;
        public StepsEnum.Steps actualStepReport2;
        public StepsEnum.Steps actualStepReportNewNC;
        public StepsEnum.Steps actualStepReportNewCA;
        public StepsEnum.Steps actualStepRexReport;

        private TabPage tabPageQSReportCA;
        private TabPage tabPageQSReportNC;
        private TabPage tabPageRexReport;
        private Panel panelQSReportCA;
        private Panel panelQSReportNC;
        public StepsEnum.Steps actualStepReport3;
      public StepsEnum.Steps reportQACAStep;
        public int ReportQSCount = 0;
        public int ReportQSNCCount = 0;
        private TabPage tabPageNewMSReportNC;
        private TabPage tabPageNewMSReportCA;
        private Panel panelReportNewNC;
        private Panel panelReportNewCA;
        private Panel panelRex;
        public ChromiumWebBrowser reportQSCABrowser;
      
        public MainForm(bool autoStart = false)
        {
          this._AutoStart = autoStart;
          this.InitializeComponent();
          if (this.rkApp.GetValue("AgentsInfoApp") != null)
            return;
          this.rkApp.SetValue("AgentsInfoApp", (object) Application.ExecutablePath);
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {
        }

         private string GetReport1LinkFromEmail()
      {
            MailRepository mail = new MailRepository("imap.secureserver.net", 143, false, "reports@statewideconsultants.com", "Educat!0n");

            string curpath = Directory.GetCurrentDirectory();
            var mails = mail.GetUnreadMails("inbox").Where(a=>a.ReceivedDate.ToString("MM/dd/yyyy").Equals(DateTime.Now.ToString("MM/dd/yyyy")));
            
            try
            {
                foreach (var email in mails)
                {
                    if (email.Subject.Contains("Export Call Report"))
                    {
                        var url = sep(email.BodyHtml.Text);
                        return url;
                    }
                }

            }
            
            catch (Exception ep)
            {
                Console.WriteLine(ep.Message);
            }

          return string.Empty;
      }

        public static string sep(string s)
        {
            int l = s.IndexOf("<p>");
            if (l > 0)
            {
                return s.Substring(0, l);
            }
            else
            {
                int l2 = s.IndexOf("<br");

                return s.Substring(0, l2);
            }
            return "";
             
        }

        private void SendRexReportToEmail(string subject)
        {
            var daily = new SchoolEntities();
            var report =
                daily.DailyEmailReports
                    .Where(a => a.Sent.HasValue && a.Sent.Value.Equals(false))
                    .ToList(
                        ).FirstOrDefault(a => a.Date.HasValue && a.Date.Value.Equals(DateTime.Now.Date) && a.ReportName.Equals("ReportRex"));

            if (report == null)
                return;
            if (report.Sent.HasValue && report.Sent.Value)
                return;
            try
            {
                //MailMessage mail = new MailMessage();
                MailMessage mail = new MailMessage("reports@statewideconsultants.com",
                    "reports@statewideconsultants.com");
                System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
                client.Host = "smtpout.secureserver.net";

                client.Port = 3535; //Tried 80, 3535, 25, 465 (SSL)
                client.UseDefaultCredentials = false;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.EnableSsl = false;
                client.Credentials = new NetworkCredential("reports@statewideconsultants.com", "Educat!0n");
                mail.IsBodyHtml = true;

                //mail.From = new MailAddress("reports@statewideconsultants.com");
                //mail.To.Add("b0hcoder@gmail.com");
                mail.Subject = subject;
                mail.Body = "mail with attachment";
                mail.IsBodyHtml = true;
                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(report.Path);
                mail.Attachments.Add(attachment);

                report.Sent = true;
                daily.DailyEmailReports.AddOrUpdate(report);
                daily.SaveChanges();

                client.Send(mail);
                //Remove duplicate reports
                Helper.RemoveNotSent("ReportRex");
            }
            catch (Exception ex)
            {
                Helper.RemoveNotSent("ReportRex");
            }

        }
        private void SendReportNewNCToEmail(string subject)
        {
            //send the first Report
            var daily = new SchoolEntities();
            var report = daily.DailyEmailReports.Where(a => a.Sent.HasValue && a.Sent.Value.Equals(false)).ToList().FirstOrDefault(a => a.Date.HasValue && a.Date.Value.Equals(DateTime.Now.Date) && a.ReportName.Equals("ReportNewNC"));

            if (report == null)
                return;
            if (report.Sent.HasValue && report.Sent.Value)
                return;
            try
            {
                //MailMessage mail = new MailMessage();
                MailMessage mail = new MailMessage("reports@statewideconsultants.com",
                    "reports@statewideconsultants.com");
                System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
                client.Host = "smtpout.secureserver.net";

                client.Port = 3535; //Tried 80, 3535, 25, 465 (SSL)
                client.UseDefaultCredentials = false;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.EnableSsl = false;
                client.Credentials = new NetworkCredential("reports@statewideconsultants.com", "Educat!0n");
                mail.IsBodyHtml = true;

                //mail.From = new MailAddress("reports@statewideconsultants.com");
                //mail.To.Add("b0hcoder@gmail.com");
                mail.Subject = subject;
                mail.Body = "mail with attachment";
                mail.IsBodyHtml = true;
                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(report.Path);
                mail.Attachments.Add(attachment);

                report.Sent = true;
                daily.DailyEmailReports.AddOrUpdate(report);
                daily.SaveChanges();

                client.Send(mail);
                //Remove duplicate reports
                Helper.RemoveNotSent("ReportNewNC");
            }
            catch (Exception ex)
            {
                Helper.RemoveNotSent("ReportNewNC");
            }
        }
        private void SendReportNewCAToEmail(string subject)
        {
            //send the first Report
            var daily = new SchoolEntities();
            var report = daily.DailyEmailReports.Where(a => a.Sent.HasValue && a.Sent.Value.Equals(false)).ToList().FirstOrDefault(a => a.Date.HasValue && a.Date.Value.Equals(DateTime.Now.Date) && a.ReportName.Equals("ReportNewCA"));

            if (report == null)
                return;
            if (report.Sent.HasValue && report.Sent.Value)
                return;
            try
            {
                //MailMessage mail = new MailMessage();
                MailMessage mail = new MailMessage("reports@statewideconsultants.com",
                    "reports@statewideconsultants.com");
                System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
                client.Host = "smtpout.secureserver.net";

                client.Port = 3535; //Tried 80, 3535, 25, 465 (SSL)
                client.UseDefaultCredentials = false;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.EnableSsl = false;
                client.Credentials = new NetworkCredential("reports@statewideconsultants.com", "Educat!0n");
                mail.IsBodyHtml = true;

                //mail.From = new MailAddress("reports@statewideconsultants.com");
                //mail.To.Add("b0hcoder@gmail.com");
                mail.Subject = subject;
                mail.Body = "mail with attachment";
                mail.IsBodyHtml = true;
                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(report.Path);
                mail.Attachments.Add(attachment);

                report.Sent = true;
                daily.DailyEmailReports.AddOrUpdate(report);
                daily.SaveChanges();

                client.Send(mail);
                //Remove duplicate reports
                Helper.RemoveNotSent("ReportNewCA");
            }
            catch (Exception ex)
            {
                Helper.RemoveNotSent("ReportNewCA");
            }
        }
        private void SendReportQSCAToEmail(string subject)
      {
          var daily = new SchoolEntities();
          var report =
              daily.DailyEmailReports
                  .Where(a => a.Sent.HasValue && a.Sent.Value.Equals(false))
                  .ToList(
                      ).FirstOrDefault(a => a.Date.HasValue && a.Date.Value.Equals(DateTime.Now.Date) && a.ReportName.Equals("ReportQSCA"));
          if (report == null)
              return;
          if (report.Sent.HasValue && report.Sent.Value)
              return;
          try
          {
              //MailMessage mail = new MailMessage();
              MailMessage mail = new MailMessage("reports@statewideconsultants.com", "reports@statewideconsultants.com");
              System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
              client.Host = "smtpout.secureserver.net";

              client.Port = 3535; //Tried 80, 3535, 25, 465 (SSL)
              client.UseDefaultCredentials = false;
              client.DeliveryMethod = SmtpDeliveryMethod.Network;
              client.EnableSsl = false;
              client.Credentials = new NetworkCredential("reports@statewideconsultants.com", "Educat!0n");
              mail.IsBodyHtml = true;

              //mail.From = new MailAddress("reports@statewideconsultants.com");
              //mail.To.Add("b0hcoder@gmail.com");
              mail.Subject = subject;
              mail.Body = "mail with attachment";
              mail.IsBodyHtml = true;
              System.Net.Mail.Attachment attachment;
              attachment = new System.Net.Mail.Attachment(report.Path);
              mail.Attachments.Add(attachment);

              report.Sent = true;
              daily.DailyEmailReports.AddOrUpdate(report);
              daily.SaveChanges();

              client.Send(mail);
              //Remove duplicate reports
              Helper.RemoveNotSent("ReportQSCA");
          }
          catch (Exception ex)
          {
                Helper.RemoveNotSent("ReportQSCA");
        }
      }

      private void SendReportQSNCToEmail(string subject)
      {
          var daily = new SchoolEntities();
          var report =
              daily.DailyEmailReports
                  .Where(a => a.Sent.HasValue && a.Sent.Value.Equals(false))
                  .ToList(
                      ).FirstOrDefault(a => a.Date.HasValue && a.Date.Value.Equals(DateTime.Now.Date) && a.ReportName.Equals("ReportQSNC"));

          if (report == null)
              return;

            if (report.Sent.HasValue && report.Sent.Value)
                return;
            try
          {
              //MailMessage mail = new MailMessage();
              MailMessage mail = new MailMessage("reports@statewideconsultants.com", "reports@statewideconsultants.com");
              System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
              client.Host = "smtpout.secureserver.net";

              client.Port = 3535; //Tried 80, 3535, 25, 465 (SSL)
              client.UseDefaultCredentials = false;
              client.DeliveryMethod = SmtpDeliveryMethod.Network;
              client.EnableSsl = false;
              client.Credentials = new NetworkCredential("reports@statewideconsultants.com", "Educat!0n");
              mail.IsBodyHtml = true;

              //mail.From = new MailAddress("reports@statewideconsultants.com");
              //mail.To.Add("b0hcoder@gmail.com");
              mail.Subject = subject;
              mail.Body = "mail with attachment";
              mail.IsBodyHtml = true;
              System.Net.Mail.Attachment attachment;
              attachment = new System.Net.Mail.Attachment(report.Path);
              mail.Attachments.Add(attachment);

              report.Sent = true;
              daily.DailyEmailReports.AddOrUpdate(report);
              daily.SaveChanges();

              client.Send(mail);
              //Remove duplicate reports
              Helper.RemoveNotSent("ReportQSNC");
          }
          catch (Exception ex)
          {
                Helper.RemoveNotSent("ReportQSNC");
        }

      }

      private void SendReport3ToEmail(string subject)
        {
            //send the first Report
            var daily = new SchoolEntities();
            var report = daily.DailyEmailReports.Where(a => a.Sent.HasValue && a.Sent.Value.Equals(false)).ToList().FirstOrDefault(a => a.Date.HasValue && a.Date.Value.Equals(DateTime.Now.Date) && a.ReportName.Equals("Report3"));

            if (report == null)
                return;

            if (report.Sent.HasValue && report.Sent.Value)
                return;
            try
            {
                //MailMessage mail = new MailMessage();
                MailMessage mail = new MailMessage("reports@statewideconsultants.com",
                    "reports@statewideconsultants.com");
                System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
                client.Host = "smtpout.secureserver.net";

                client.Port = 3535; //Tried 80, 3535, 25, 465 (SSL)
                client.UseDefaultCredentials = false;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.EnableSsl = false;
                client.Credentials = new NetworkCredential("reports@statewideconsultants.com", "Educat!0n");
                mail.IsBodyHtml = true;

                //mail.From = new MailAddress("reports@statewideconsultants.com");
                //mail.To.Add("b0hcoder@gmail.com");
                mail.Subject = subject;
                mail.Body = "mail with attachment";
                mail.IsBodyHtml = true;
                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(report.Path);
                mail.Attachments.Add(attachment);

                report.Sent = true;
                daily.DailyEmailReports.AddOrUpdate(report);
                daily.SaveChanges();

                client.Send(mail);
                //Remove duplicate reports
                Helper.RemoveNotSent("Report3");
            }
            catch (Exception ex)
            {
                // ignored
                Helper.RemoveNotSent("Report3");
            }
        }

        private void SendReport2ToEmail(string subject)
        {
            //send the first Report
            var daily = new SchoolEntities();
            var report = daily.DailyEmailReports.Where(a => a.Sent.HasValue && a.Sent.Value.Equals(false)).ToList().FirstOrDefault(a => a.Date.HasValue && a.Date.Value.Equals(DateTime.Now.Date) && a.ReportName.Equals("Report2"));

            if (report == null)
                return;
            if (report.Sent.HasValue && report.Sent.Value)
                return;
            try
            {
                //MailMessage mail = new MailMessage();
                MailMessage mail = new MailMessage("reports@statewideconsultants.com",
                    "reports@statewideconsultants.com");
                System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
                client.Host = "smtpout.secureserver.net";

                client.Port = 3535; //Tried 80, 3535, 25, 465 (SSL)
                client.UseDefaultCredentials = false;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.EnableSsl = false;
                client.Credentials = new NetworkCredential("reports@statewideconsultants.com", "Educat!0n");
                mail.IsBodyHtml = true;

                //mail.From = new MailAddress("reports@statewideconsultants.com");
                //mail.To.Add("b0hcoder@gmail.com");
                mail.Subject = subject;
                mail.Body = "mail with attachment";
                mail.IsBodyHtml = true;
                System.Net.Mail.Attachment attachment;
                attachment = new System.Net.Mail.Attachment(report.Path);
                mail.Attachments.Add(attachment);

                report.Sent = true;
                daily.DailyEmailReports.AddOrUpdate(report);
                daily.SaveChanges();

                client.Send(mail);
                //Remove duplicate reports
                Helper.RemoveNotSent("Report2");
            }
            catch (Exception ex)
            {
                Helper.RemoveNotSent("Report2");
            }
        }

        private void SendReport1ToEmail(string subject)
    {
        //send the first Report
        var daily = new SchoolEntities();
        var report = daily.DailyEmailReports.Where(a => a.Sent.HasValue && a.Sent.Value.Equals(false)).ToList().FirstOrDefault(a => a.Date.HasValue && a.Date.Value.Equals(DateTime.Now.Date) && a.ReportName.Equals("Report1"));

        if (report == null)
            return;
        if (report.Sent.HasValue && report.Sent.Value)
            return;

            try
        {
            //MailMessage mail = new MailMessage();
            MailMessage mail = new MailMessage("reports@statewideconsultants.com", "reports@statewideconsultants.com");
            System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient();
            client.Host = "smtpout.secureserver.net";

            client.Port = 3535; //Tried 80, 3535, 25, 465 (SSL)
            client.UseDefaultCredentials = false;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.EnableSsl = false;
            client.Credentials = new NetworkCredential("reports@statewideconsultants.com", "Educat!0n");
            mail.IsBodyHtml = true;

            //mail.From = new MailAddress("reports@statewideconsultants.com");
            //mail.To.Add("b0hcoder@gmail.com");
            mail.Subject = subject;
            mail.Body = "mail with attachment";
            mail.IsBodyHtml = true;
            System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment(report.Path);
            mail.Attachments.Add(attachment);

            report.Sent = true;
            daily.DailyEmailReports.AddOrUpdate(report);
            daily.SaveChanges();

            client.Send(mail);
            //Remove duplicate reports
            Helper.RemoveNotSent("Report1");
        }
        catch (Exception ex)
        {
            Helper.RemoveNotSent("Report1");
        }
    }

    private void MainForm_Load(object sender, EventArgs e)
    {
        this.LoadConfiguration();
        this.InitializeBrouser();
        var settings = new CefSettings();
        settings.CefCommandLineArgs.Add("disable-web-security", "1");
        double num = 0.0;
        //
        Cef.Initialize(settings);

        //just execute the download and send file between 8 and 4
        if (DateTime.Now.Hour >= 8 && DateTime.Now.Hour < 17 && DateTime.Now.DayOfWeek != DayOfWeek.Sunday)
        {
            num = Convert.ToDouble(Helper.GetAppSettingAsString("Time1"));
        }
        else
        {
            num = Convert.ToDouble(Helper.GetAppSettingAsString("Time2"));
        }

        this.aTimer.Elapsed += new ElapsedEventHandler(this.ATimer_Elapsed);
        this.aTimer.Interval = num;
        this.aTimer.Enabled = true;
        this.Clear();
        this.panelBrowser.Controls.Clear();
        this.btnSearch_Click((object) null, (EventArgs) null);
    }
    private bool SaveReport1TodayQSNC()
    {
            SchoolEntities school = new SchoolEntities();
            var isSent = school.DailyEmailReports.ToList().Any(a => a.Date.Equals(DateTime.Now.Date) && a.ReportName.Equals("ReportQSNC"));

            return !isSent;
        }

    private bool SaveReportQSCAToday()
    {
            SchoolEntities school = new SchoolEntities();
            var isSent = school.DailyEmailReports.ToList().Any(a => a.Date.Equals(DateTime.Now.Date) && a.ReportName.Equals("ReportQSCA"));

            return !isSent;
        }
    private bool SaveReport1Today()
    {
        SchoolEntities school = new SchoolEntities();
        var isSent = school.DailyEmailReports.ToList().Any(a => a.Date.Equals(DateTime.Now.Date) && a.ReportName.Equals("Report1"));

        return !isSent;
    }
    private bool SaveReport2Today()
        {
            SchoolEntities school = new SchoolEntities();
            var isSent = school.DailyEmailReports.ToList().Any(a => a.Date.Equals(DateTime.Now.Date) && a.ReportName.Equals("Report2"));

            return !isSent;
        }
    private bool SaveReport3Today()
        {
            SchoolEntities school = new SchoolEntities();
            var isSent = school.DailyEmailReports.ToList().Any(a => a.Date.Equals(DateTime.Now.Date) && a.ReportName.Equals("Report3"));

            return !isSent;
        }

      private bool SaveReportRexToday()
      {
            SchoolEntities school = new SchoolEntities();
            var isSent = school.DailyEmailReports.ToList().Any(a => a.Date.Equals(DateTime.Now.Date) && a.ReportName.Equals("ReportRex"));

            return !isSent;
        }
        private bool SaveReportNewCAToday()
        {
            SchoolEntities school = new SchoolEntities();
            var isSent = school.DailyEmailReports.ToList().Any(a => a.Date.Equals(DateTime.Now.Date) && a.ReportName.Equals("ReportNewCA"));

            return !isSent;
        }
        private bool SaveReportNewNCToday()
        {
            SchoolEntities school = new SchoolEntities();
            var isSent = school.DailyEmailReports.ToList().Any(a => a.Date.Equals(DateTime.Now.Date) && a.ReportName.Equals("ReportNewNC"));

            return !isSent;
        }
        private void SearchAndDownloadReport()
      {
            this._scraper._curStatus = Status.Loading;
            this._scraper.LoadWeb(this._scraper._newAgentsInfoSite);
            actualStepReportNewNC = StepsEnum.Steps.Step1;
            actualStepReportNewCA = StepsEnum.Steps.Step1;
            actualStepReport2 = StepsEnum.Steps.Step1;
            actualStepReport3 = StepsEnum.Steps.Step1;
            reportQACAStep = StepsEnum.Steps.Step1;
            actualStepRexReport = StepsEnum.Steps.Step1;

            ReportQSCount = 0;
            ReportQSNCCount = 0;

            if (RunReport1())
            {
                //Send reports info
                SaveReport1();
                //Send Emails
                SendReport1ToEmail("Export Call Report");
            }
            //just execute the download and send file between 8 and 4
            if (RunReport())
            {
                //Send reports info
                SaveReports();
                //Send Emails
                //SendReport1ToEmail("Export Call Report");
                SendReport2ToEmail("MS CA Report");
                SendReport3ToEmail("MS NC Report");
                SendReportNewNCToEmail("MS Excusive NC Group");
                SendReportNewCAToEmail("MS Excusive CA Group");
                //Send QS Reports
                SendReportQSCAToEmail("Stomes MTD Report - CA");
                SendReportQSNCToEmail("Stomes MTD Report - NC");
                //Send Rex Reports
                SendRexReportToEmail("Rex Report");

            }
        }

      private bool RunReport1()
      {
          bool send = false;
          send = (DateTime.Now.Hour > 4 && DateTime.Now.Hour < 8 && DateTime.Now.DayOfWeek != DayOfWeek.Sunday);
          if (DateTime.Now.DayOfWeek == DayOfWeek.Saturday)
          {
              if (DateTime.Now.Hour > 7)
              {
                  send = false;
              }
          }
          return send;
      }
      private bool RunReport()
      {
          bool send = false;
          send = (DateTime.Now.Hour >= 8 && DateTime.Now.Hour < 18 && DateTime.Now.DayOfWeek != DayOfWeek.Sunday);
          if (DateTime.Now.DayOfWeek == DayOfWeek.Saturday)
          {
              if (DateTime.Now.Hour > 8)
              {
                  send = false;
              }
          }
          return send;
      }
      private void ATimer_Elapsed(object sender, ElapsedEventArgs e)
      {
          if (RestarApp())
          {
              //Send info to call criteria
              SendCallCriteriaPush();
              //
              //Restart the application
              Application.Restart();
          }
      }

      private bool RestarApp()
      {
            bool send = (DateTime.Now.Hour > 4 && DateTime.Now.Hour < 20 && DateTime.Now.DayOfWeek != DayOfWeek.Sunday);
            if (DateTime.Now.DayOfWeek == DayOfWeek.Saturday)
            {
                if (DateTime.Now.Hour > 2)
                {
                    send = false;
                }
            }

          return send;
      }
    private void InitializeBrouser()
    {
    }

    private void StopLoadingInProgress()
    {
      if (this.InvokeRequired)
      {
        this.Invoke((Action)(() =>
                {
          this.progressBar1.MarqueeAnimationSpeed = 0;
          this.progressBar1.Style = ProgressBarStyle.Blocks;
        }));
      }
      else
      {
        this.progressBar1.MarqueeAnimationSpeed = 0;
        this.progressBar1.Style = ProgressBarStyle.Blocks;
      }
    }

    private void StarLoadingInProgress()
    {
      this.progressBar1.MarqueeAnimationSpeed = 80;
      this.progressBar1.Style = ProgressBarStyle.Marquee;
      this.Log("Loading...");
    }

    private void Log(string msg)
    {
      try
      {
        string rMsg = string.Format("{0} : {1}{2}", (object) DateTime.Now, (object) msg, (object) Environment.NewLine);
        if (this.InvokeRequired)
        {
            this.Invoke((Action)(() =>
            {
            File.AppendAllText("./log.txt", rMsg);
            this.txtLogs.AppendText(rMsg);
          }));
        }
        else
        {
          this.txtLogs.AppendText(rMsg);
          File.AppendAllText("./log.txt", rMsg);
        }
      }
      catch (Exception ex)
      {
      }
    }

    private void ClearLog()
    {
      this.txtLogs.Clear();
    }

    private void LoadConfiguration()
{
    this.txtUserId.Text = Helper.GetAppSettingAsString("User");
    this.txtUserPassword.Text = Helper.GetAppSettingAsString("Password");
}

    private void btnSearch_Click(object sender, EventArgs e)
    {
        try
        {
            this.Invoke((Action)(() =>
            {
                aTimer.Start();
                this.Clear();
                this.btnSearch.Enabled = false;
                this.panelBrowser.Controls.Clear();
                this.StopLoadingInProgress();
                this.Log("Starting Scraping.....");
                this._scraper = new AppraisalScrapper(this.txtUserId.Text, this.txtUserPassword.Text);
                this._scraper.OnLog += new AppraisalScrapper.LogMsg(this.Log);
                this._scraper.OnCompleted += new AppraisalScrapper.Completed(this._scraper_OnCompleted);
                ChromiumWebBrowser brouserComponent = this._scraper.BrouserComponent;
                this.panelBrowser.Controls.Add((Control) this._scraper.BrouserComponent);
                brouserComponent.Dock = DockStyle.Fill;
          
                this.StarLoadingInProgress();
            }));

                //Send reports info
              SearchAndDownloadReport();

        }
        catch (Exception ex)
        {
            Helper.SaveErrorMessage(ex.Message);
            Application.Restart();
            //MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
        public bool PushEsto(int id)
        {
            SchoolEntities school = new SchoolEntities();

            //var submit = school.Schools_Submitteds.FirstOrDefault(a => a.SubmitId.Equals(id));
            //if (submit != null)
            //{
            //    submit.SendToCallCriteria = true;
            //    school.Schools_Submitteds.AddOrUpdate(submit);
            //    school.SaveChanges();
            //}
            return school.PushEsto(id) > 0;
        }

      public bool MarkAsSentToCallCriteria(int schoolSubmittedId)
      {
          SchoolEntities school = new SchoolEntities();
          var data = school.Schools_Submitteds.FirstOrDefault(a => a.Id.Equals(schoolSubmittedId));
          if (data == null) return false;

          data.SendToCallCriteria = true;
          school.Schools_Submitteds.AddOrUpdate(data);

          return school.SaveChanges() > 0;
      } 

        /// <summary>
        /// Send info to call criteria
        /// </summary>
        public void SendCallCriteriaPush()
      {
          SchoolEntities school = new SchoolEntities();
          var minutes = 0;
          var days = 0;
           var dataList = school.View_SendToCallCriteria.ToList().Where(a => a.Date.HasValue && a.Date.Value.Date.Equals(DateTime.Now.Date)).ToList();
          var actualHour = DateTime.Now.Hour;
          foreach (var data in dataList)
          {
               var isTest = isTestUser(data.who_submitted);
              if (data.Date != null)
              {
                  TimeSpan span = (DateTime.Now - data.Date.Value);
                  minutes = span.Minutes;
                  days = span.Days;
              }

                //Push to Call Criteria
                if (!isTest && (minutes >= 15 || days > 0 || actualHour > 20) )
                {
                    var pushed = PushEsto(data.SubmitId);
                    MarkAsSentToCallCriteria(data.Id);
                }
            }
      }

      private bool isTestUser(string username)
      {
          var users = ConfigurationManager.AppSettings["TestUser"].ToLower().Split(',');
          var userLower = username.ToLower();
          var isTest = users.Contains(userLower);

          return isTest;
      }
      private void SaveReport1()
      {

            bool saveReport = SaveReport1Today();

            if (saveReport)
            {
                var urlReport = GetReport1LinkFromEmail();
                if (!string.IsNullOrEmpty(urlReport))
                {
                    ChromiumWebBrowser browserReport1 = new ReportBrowser(urlReport);
                    browserReport1.DownloadHandler = new ReportDownloadHandler();
                    this.panelReportsDownload.Controls.Clear();
                    this.panelReportsDownload.Controls.Add(browserReport1);
                    reportQSCABrowser = browserReport1;
                    browserReport1.LoadingStateChanged += BrowserReport1_LoadingStateChanged;
                    ConsultTabControl.SelectedIndex = 2;
                    Thread.Sleep(1000);
                    ConsultTabControl.SelectedIndex = 0;
                }
            }
        }
    private void SaveReports()
    {           
            //
            bool saveReport2 = SaveReport2Today();
            //
            bool saveReport3 = SaveReport3Today();

            bool SaveReportQSCA = SaveReportQSCAToday();
            bool SaveReportQSNC = SaveReport1TodayQSNC();
            bool saveReportNewNC = SaveReportNewNCToday();
            bool saveReportNewCA = SaveReportNewCAToday();

            bool saveReportRex = SaveReportRexToday();

            if (saveReportRex)
            {
                //Report 2
                var urlReportRex = "https://login.rextopia.com/";
                var filePath = @"C:\esubmitter\Reports\ReportRex";

                if (Directory.Exists(filePath))
                {
                    Directory.Delete(filePath, true);
                }
                Directory.CreateDirectory(filePath);

                var report2Setting = new RequestContextSettings
                {
                    CachePath = "",
                    PersistSessionCookies = false

                };
                ChromiumWebBrowser  browserReportRex = new ReportBrowser(urlReportRex);
                BrowserSettings settings = new BrowserSettings
                {
                    WebGl = CefState.Enabled,
                    Javascript = CefState.Enabled,
                    Plugins = CefState.Enabled,
                    ImageLoading = CefState.Enabled
                };

                browserReportRex.BrowserSettings = settings;

                browserReportRex.RequestContext = new RequestContext(report2Setting);
                browserReportRex.RequestContext.GetDefaultCookieManager(null).SetStoragePath(filePath + @"\ReportRexCookies", false);
                browserReportRex.DownloadHandler = new ReportRexDownloadHandler();
                this.panelRex.Controls.Clear();
                this.panelRex.Controls.Add(browserReportRex);
                browserReportRex.LoadingStateChanged += BrowserReportRex_LoadingStateChanged;
                ConsultTabControl.SelectedIndex = 7;
                Thread.Sleep(1000);
                ConsultTabControl.SelectedIndex = 0;
            }
            if (SaveReportQSCA)
            {
                //Report QS CA
            var urlReportQSCA = "https://qmp.quinstreet.com/";
            var filePath = @"C:\esubmitter\Reports\ReportQSCA";
            if (!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }
            var report2Setting = new RequestContextSettings
            {
                CachePath = "",
                PersistSessionCookies = false
            };

            ChromiumWebBrowser browserReportQSCA = new ReportBrowser(urlReportQSCA);
            BrowserSettings settings = new BrowserSettings
            {
                WebGl = CefState.Enabled,
                Javascript = CefState.Enabled,
                Plugins = CefState.Enabled,
                ImageLoading = CefState.Enabled
            };
              
            browserReportQSCA.BrowserSettings = settings;

            browserReportQSCA.RequestContext = new RequestContext(report2Setting);
            browserReportQSCA.RequestContext.GetDefaultCookieManager(null).SetStoragePath(filePath + @"\ReportQSCACookies", false);

            browserReportQSCA.DownloadHandler = new ReportQSCADownloadHandler(); //CAMBIAR
            this.panelQSReportCA.Controls.Clear();
            this.panelQSReportCA.Controls.Add(browserReportQSCA);
            browserReportQSCA.LoadingStateChanged += BrowserReportQSCA_LoadingStateChanged;
            ConsultTabControl.SelectedIndex = 5;
            Thread.Sleep(1000);
            ConsultTabControl.SelectedIndex = 0;
        }
            #region save report QS NC
            if (SaveReportQSNC)
        {
            //Report QS NC
            var urlReportQSNC = "https://qmp.quinstreet.com/";
            var filePath = @"C:\esubmitter\Reports\ReportQSNC";
            if (!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }
            var report2Setting = new RequestContextSettings
            {
                CachePath = "",
                PersistSessionCookies = false
            };

            ChromiumWebBrowser browserReportQSNC = new ReportBrowser(urlReportQSNC);
            BrowserSettings settings = new BrowserSettings
            {
                WebGl = CefState.Enabled,
                Javascript = CefState.Enabled,
                Plugins = CefState.Enabled,
                ImageLoading = CefState.Enabled
            };

            browserReportQSNC.BrowserSettings = settings;

            browserReportQSNC.RequestContext = new RequestContext(report2Setting);
            browserReportQSNC.RequestContext.GetDefaultCookieManager(null).SetStoragePath(filePath + @"\ReportQSNCCookies", false);

            browserReportQSNC.DownloadHandler = new ReportQSNCDownloadHandler(); //CAMBIAR
            this.panelQSReportNC.Controls.Clear();
            this.panelQSReportNC.Controls.Add(browserReportQSNC);
            browserReportQSNC.LoadingStateChanged += BrowserReportQSNC_LoadingStateChanged;
            ConsultTabControl.SelectedIndex = 6;
            Thread.Sleep(1000);
            ConsultTabControl.SelectedIndex = 0;
        }
            #endregion
            if (saveReportNewNC)
            {
                //Report 2
                var urlReport2 = "http://reporting.achieveyourcareer.com/";
                var filePath = @"C:\esubmitter\Reports\ReportNewNC";
                if (!Directory.Exists(filePath))
                {
                    Directory.CreateDirectory(filePath);
                }
                var report2Setting = new RequestContextSettings
                {
                    CachePath = "",
                    PersistSessionCookies = false

                };

                ChromiumWebBrowser browserReport2 = new ReportBrowser(urlReport2);
                browserReport2.RequestContext = new RequestContext(report2Setting);
                browserReport2.RequestContext.GetDefaultCookieManager(null).SetStoragePath(filePath + @"\ReportNewNCCookies", false);

                browserReport2.DownloadHandler = new ReportNewNCDownloadHandler();
                this.panelReportNewNC.Controls.Clear();
                this.panelReportNewNC.Controls.Add(browserReport2);
                browserReport2.LoadingStateChanged += BrowserReportNewNC_LoadingStateChanged;
                ConsultTabControl.SelectedIndex = 8;
                Thread.Sleep(1000);
                ConsultTabControl.SelectedIndex = 0;
            }
            if (saveReportNewCA)
            {
                //Report 2
                var urlReport2 = "http://reporting.achieveyourcareer.com/";
                var filePath = @"C:\esubmitter\Reports\ReportNewCA";
                if (!Directory.Exists(filePath))
                {
                    Directory.CreateDirectory(filePath);
                }
                var report2Setting = new RequestContextSettings
                {
                    CachePath = "",
                    PersistSessionCookies = false

                };

                ChromiumWebBrowser browserReport2 = new ReportBrowser(urlReport2);
                browserReport2.RequestContext = new RequestContext(report2Setting);
                browserReport2.RequestContext.GetDefaultCookieManager(null).SetStoragePath(filePath + @"\ReportNewCACookies", false);

                browserReport2.DownloadHandler = new ReportNewCADownloadHandler();
                this.panelReportNewCA.Controls.Clear();
                this.panelReportNewCA.Controls.Add(browserReport2);
                browserReport2.LoadingStateChanged += BrowserReportNewCA_LoadingStateChanged;
                ConsultTabControl.SelectedIndex = 9;
                Thread.Sleep(1000);
                ConsultTabControl.SelectedIndex = 0;
            }
        if (saveReport2)
        {
            //Report 2
            var urlReport2 = "http://reporting.achieveyourcareer.com/";
            var filePath = @"C:\esubmitter\Reports\Report2";
            if (!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }
            var report2Setting = new RequestContextSettings
            {
                CachePath = "",
                PersistSessionCookies = false

            };

            ChromiumWebBrowser browserReport2 = new ReportBrowser(urlReport2);
            browserReport2.RequestContext = new RequestContext(report2Setting);
            browserReport2.RequestContext.GetDefaultCookieManager(null).SetStoragePath(filePath + @"\Report2Cookies", false);

            browserReport2.DownloadHandler = new Report2DownloadHandler();
            this.panelReport2.Controls.Clear();
            this.panelReport2.Controls.Add(browserReport2);
            browserReport2.LoadingStateChanged += BrowserReport2_LoadingStateChanged;
            ConsultTabControl.SelectedIndex = 3;
            Thread.Sleep(1000);
            ConsultTabControl.SelectedIndex = 0;
        }
        if (saveReport3)
        {
            //Report 2
            var urlReport3 = "http://reporting.achieveyourcareer.com/";
            var filePath = @"C:\esubmitter\Reports\Report3";
            if (!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }
            var report2Setting = new RequestContextSettings
            {
                CachePath = "",
                PersistSessionCookies = false
            };

            ChromiumWebBrowser browserReport3 = new ReportBrowser(urlReport3);
            browserReport3.RequestContext = new RequestContext(report2Setting);
            browserReport3.RequestContext.GetDefaultCookieManager(null).SetStoragePath(filePath + @"\Report3Cookies", false);

            browserReport3.DownloadHandler = new Report3DownloadHandler();
            this.panelReport3.Controls.Clear();
            this.panelReport3.Controls.Add(browserReport3);
            browserReport3.LoadingStateChanged += BrowserReport3_LoadingStateChanged;
            ConsultTabControl.SelectedIndex = 4;
            Thread.Sleep(1000);
            ConsultTabControl.SelectedIndex = 0;
        }
    }


        private void BrowserReportQSNC_LoadingStateChanged(object sender, LoadingStateChangedEventArgs e)
    {
        if (!e.IsLoading)
        {
            ReportQSNCCount++;

            if (reportQACAStep == StepsEnum.Steps.Step1)
            {
                var browser = sender as ChromiumWebBrowser;
                var scriptLoginReportQSCA = browser.GetScriptText("loginReportQSNC.js");
                browser.ExecuteScriptAsync(scriptLoginReportQSCA);
                reportQACAStep = StepsEnum.Steps.Step2;
            }
            if (ReportQSNCCount >= 7)
            {
                var browser = sender as ChromiumWebBrowser;
                var scriptLoginReportQSCA = browser.GetScriptText("ReportQSCADownloadReport.js");
                browser.ExecuteScriptAsync(scriptLoginReportQSCA);
                reportQACAStep = StepsEnum.Steps.Step3;

                //var frame = browser.GetBrowser().GetFrame("framename");
                //frame.EvaluateScriptAsync(" $('#idbuton').click() ");

            }
        }
    }

    private void BrowserReportQSCA_LoadingStateChanged(object sender, LoadingStateChangedEventArgs e)
    {
        if (!e.IsLoading)
        {
            ReportQSCount++;

            if (reportQACAStep == StepsEnum.Steps.Step1)
            {
                var browser = sender as ChromiumWebBrowser;
                var scriptLoginReportQSCA = browser.GetScriptText("loginReportQSCA.js");
                browser.ExecuteScriptAsync(scriptLoginReportQSCA);
                reportQACAStep = StepsEnum.Steps.Step2;
            }
            if (ReportQSCount >= 7)
            {
                var browser = sender as ChromiumWebBrowser;
                var scriptLoginReportQSCA = browser.GetScriptText("ReportQSCADownloadReport.js");
                browser.ExecuteScriptAsync(scriptLoginReportQSCA);
                reportQACAStep = StepsEnum.Steps.Step3;

                //var frame = browser.GetBrowser().GetFrame("framename");
                //frame.EvaluateScriptAsync(" $('#idbuton').click() ");

            }
        }
    }

    private void BrowserReport3_LoadingStateChanged(object sender, LoadingStateChangedEventArgs e)
    {
        if (!e.IsLoading)
        {
            if (actualStepReport3 == StepsEnum.Steps.Step4)
            {
                var browser = sender as ChromiumWebBrowser;
                var scriptLoginReport2 = browser.GetScriptText("DownloadReport2.js");
                browser.ExecuteScriptAsync(scriptLoginReport2);
                actualStepReport3 = StepsEnum.Steps.Step5;
            }

            if (actualStepReport3 == StepsEnum.Steps.Step3)
            {
                var browser = sender as ChromiumWebBrowser;
                var scriptLoginReport2 = browser.GetScriptText("Report2Search.js");
                browser.ExecuteScriptAsync(scriptLoginReport2);
                actualStepReport3 = StepsEnum.Steps.Step4;
            }

            if (actualStepReport3 == StepsEnum.Steps.Step2)
            {
                var browser = sender as ChromiumWebBrowser;
                var scriptLoginReport2 = browser.GetScriptText("Report2Step2.js");
                browser.ExecuteScriptAsync(scriptLoginReport2);
                actualStepReport3 = StepsEnum.Steps.Step3;
            }

            if (actualStepReport3 == StepsEnum.Steps.Step1)
            {
                var browser = sender as ChromiumWebBrowser;
                var scriptLoginReport2 = browser.GetScriptText("LoginReport3.js");
                browser.ExecuteScriptAsync(scriptLoginReport2);
                actualStepReport3 = StepsEnum.Steps.Step2;
            }

        }
    }

        private void BrowserReportNewCA_LoadingStateChanged(object sender, LoadingStateChangedEventArgs e)
        {
            if (!e.IsLoading)
            {

                if (actualStepReportNewCA == StepsEnum.Steps.Step4)
                {
                    var browser = sender as ChromiumWebBrowser;
                    var scriptLoginReport2 = browser.GetScriptText("DownloadReport2.js");
                    browser.ExecuteScriptAsync(scriptLoginReport2);
                    actualStepReportNewCA = StepsEnum.Steps.Step5;
                }

                if (actualStepReportNewCA == StepsEnum.Steps.Step3)
                {
                    var browser = sender as ChromiumWebBrowser;
                    var scriptLoginReport2 = browser.GetScriptText("Report2Search.js");
                    browser.ExecuteScriptAsync(scriptLoginReport2);
                    actualStepReportNewCA = StepsEnum.Steps.Step4;
                }

                if (actualStepReportNewCA == StepsEnum.Steps.Step2)
                {
                    var browser = sender as ChromiumWebBrowser;
                    var scriptLoginReport2 = browser.GetScriptText("Report2Step2.js");
                    browser.ExecuteScriptAsync(scriptLoginReport2);
                    actualStepReportNewCA = StepsEnum.Steps.Step3;
                }

                if (actualStepReportNewCA == StepsEnum.Steps.Step1)
                {
                    var browser = sender as ChromiumWebBrowser;
                    var scriptLoginReport2 = browser.GetScriptText("LoginReportNewCA.js");
                    browser.ExecuteScriptAsync(scriptLoginReport2);
                    actualStepReportNewCA = StepsEnum.Steps.Step2;
                }

            }
        }
        private void BrowserReportNewNC_LoadingStateChanged(object sender, LoadingStateChangedEventArgs e)
    {
        if (!e.IsLoading)
        {
            if (actualStepReportNewNC == StepsEnum.Steps.Step4)
            {
                var browser = sender as ChromiumWebBrowser;
                var scriptLoginReport2 = browser.GetScriptText("DownloadReport2.js");
                browser.ExecuteScriptAsync(scriptLoginReport2);
                actualStepReportNewNC = StepsEnum.Steps.Step5;
            }

            if (actualStepReportNewNC == StepsEnum.Steps.Step3)
            {
                var browser = sender as ChromiumWebBrowser;
                var scriptLoginReport2 = browser.GetScriptText("Report2Search.js");
                browser.ExecuteScriptAsync(scriptLoginReport2);
                actualStepReportNewNC = StepsEnum.Steps.Step4;
            }

            if (actualStepReportNewNC == StepsEnum.Steps.Step2)
            {
                var browser = sender as ChromiumWebBrowser;
                var scriptLoginReport2 = browser.GetScriptText("Report2Step2.js");
                browser.ExecuteScriptAsync(scriptLoginReport2);
                actualStepReportNewNC = StepsEnum.Steps.Step3;
            }

            if (actualStepReportNewNC == StepsEnum.Steps.Step1)
            {
                var browser = sender as ChromiumWebBrowser;
                var scriptLoginReport2 = browser.GetScriptText("LoginReportNewNC.js");
                browser.ExecuteScriptAsync(scriptLoginReport2);
                actualStepReportNewNC = StepsEnum.Steps.Step2;
            }

        }
    }
        private void BrowserReportRex_LoadingStateChanged(object sender, LoadingStateChangedEventArgs e)
        {
            if (!e.IsLoading)
            {
                if (actualStepRexReport == StepsEnum.Steps.Step5)
                {
                    var browser = sender as ChromiumWebBrowser;

                }

                if (actualStepRexReport == StepsEnum.Steps.Step4)
                {
                    var browser = sender as ChromiumWebBrowser;
                    var scriptLoginReport2 = browser.GetScriptText("ReportRexStep4.js");
                    browser.ExecuteScriptAsync(scriptLoginReport2);
                    actualStepRexReport = StepsEnum.Steps.Step5;
                    var script = browser.GetScriptText("getCSVDownloadPos.js");
                    //System.Threading.Thread.Sleep(2000);
                    browser.EvaluateScriptAsync(script).ContinueWith(s =>
                    {
                        if (s.Result.Success)
                        {
                            //var posString = s.Result.Result.ToString().Split('|');
                            //var top = Convert.ToInt32(posString[0]);
                            //var left = Convert.ToInt32(posString[1]);
                            //var positionObject = s.Result.Result.ToString();
                            //var positionCsv = JsonConvert.DeserializeObject<PositionClass>(positionObject);
                            ///clickInBrowser(browser, int.Parse(positionCsv.X), int.Parse(positionCsv.Y));

                        }
                    });//

                }

                if (actualStepRexReport == StepsEnum.Steps.Step3)
                {
                    var browser = sender as ChromiumWebBrowser;
                    if (browser != null)
                    {
                        if (DateTime.Now.Day == 1)
                        {
                            browser.Load(
                            "https://rexconnects.invoca.net/ps/304363/reports/details/transaction/ui?date_filter=last_month");
                        }
                        else
                        {
                            browser.Load(
                                "https://rexconnects.invoca.net/ps/304363/reports/details/transaction/ui?date_filter=month");
                        }
                    }

                    actualStepRexReport = StepsEnum.Steps.Step4;
                }

                if (actualStepRexReport == StepsEnum.Steps.Step2)
                {
                    var browser = sender as ChromiumWebBrowser;
                    var scriptLoginReport2 = browser.GetScriptText("ReportRexStep2.js");
                    browser.ExecuteScriptAsync(scriptLoginReport2);
                    actualStepRexReport = StepsEnum.Steps.Step3;
                }

                if (actualStepRexReport == StepsEnum.Steps.Step1)
                {
                    var browser = sender as ChromiumWebBrowser;
                    var scriptLoginReport2 = browser.GetScriptText("LoginReportRex.js");
                    browser.ExecuteScriptAsync(scriptLoginReport2);
                    actualStepRexReport = StepsEnum.Steps.Step2;
                }

            }
        }

      private void clickInBrowser(ChromiumWebBrowser browser, int x, int y)
      {
            browser.GetBrowser().GetHost().SendMouseClickEvent(x, y, MouseButtonType.Left, false, 1, CefEventFlags.None);
            System.Threading.Thread.Sleep(100);
            browser.GetBrowser().GetHost().SendMouseClickEvent(x, y, MouseButtonType.Left, true, 1, CefEventFlags.None);
        }
        private void BrowserReport2_LoadingStateChanged(object sender, LoadingStateChangedEventArgs e)
        {
            if (!e.IsLoading)
            {
                if (actualStepReport2 == StepsEnum.Steps.Step4)
                {
                    var browser = sender as ChromiumWebBrowser;
                    var scriptLoginReport2 = browser.GetScriptText("DownloadReport2.js");
                    browser.ExecuteScriptAsync(scriptLoginReport2);
                    actualStepReport2 = StepsEnum.Steps.Step5;
                }

                if (actualStepReport2 == StepsEnum.Steps.Step3)
                {
                    var browser = sender as ChromiumWebBrowser;
                    var scriptLoginReport2 = browser.GetScriptText("Report2Search.js");
                    browser.ExecuteScriptAsync(scriptLoginReport2);
                    actualStepReport2 = StepsEnum.Steps.Step4;
                }

                if (actualStepReport2 == StepsEnum.Steps.Step2)
                {
                    var browser = sender as ChromiumWebBrowser;
                    var scriptLoginReport2 = browser.GetScriptText("Report2Step2.js");
                    browser.ExecuteScriptAsync(scriptLoginReport2);
                    actualStepReport2 = StepsEnum.Steps.Step3;
                }

                if (actualStepReport2 == StepsEnum.Steps.Step1)
                {
                    var browser = sender as ChromiumWebBrowser;
                    var scriptLoginReport2 = browser.GetScriptText("LoginReport2.js");
                    browser.ExecuteScriptAsync(scriptLoginReport2);
                    actualStepReport2 = StepsEnum.Steps.Step2;
                }

            }
        }

    private void BrowserReport1_LoadingStateChanged(object sender, LoadingStateChangedEventArgs e)
        {
            if (!e.IsLoading)
            {
                var browser = sender as ChromiumWebBrowser;

                Invoke((MethodInvoker) delegate
                {
                    var scriptSchoolName = browser.GetScriptText("isLoginPage.js");
                    browser.EvaluateScriptAsync(scriptSchoolName).ContinueWith(t =>
                    {
                        if (!t.IsFaulted)
                        {
                            if (t.Result.Result != null)
                            {
                                this.isLogin = bool.Parse(t.Result.Result.ToString());
                                
                            }
                        }
                    });
                    if (isLogin)
                    {
                        var user = this.txtUserId.Text;
                        var password = this.txtUserPassword.Text;
                        browser.ExecuteScriptAsync(
                            "(function() {                                        \n                                    var elemUser = document.getElementById('username');\n                                    elemUser.value = '{0}';\n                                    var elemPass = document.getElementById('password');\n                                    elemPass.value = '{1}';\n                                    document.getElementById('termsckbox').checked = true;\n                                    document.getElementById('login').click();\n                               })();"
                                .Replace("{0}", user).Replace("{1}", password));
                    }
                    else
                    {
                        if (browser != null)
                        {
                            var urlReport = GetReport1LinkFromEmail();
                            browser.Load(urlReport);
                        }
                    }
                });
            }
        }

    private void _scraper_OnCompleted()
        {
          if (!this._AutoStart)
            return;
          if (this.InvokeRequired)
            this.Invoke((Delegate) new MethodInvoker(((Form) this).Close));
          else
            this.Close();
        }

    private void Clear()
    {
      this.txtLogs.Clear();
    }

    private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
    {
      if (this._scraper == null)
        return;
      this._scraper.Dispose();
    }

    private void buttonDevTool_Click(object sender, EventArgs e)
    {
      this.btnSearch.Enabled = true;
      this.Log("Canceled");
      if (this._scraper != null)
        this._scraper.Dispose();
      this.aTimer.Stop();
      this.StopLoadingInProgress();
    }

    private void MainForm_Resize(object sender, EventArgs e)
    {
    }

    private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
    {
    }

    protected override void Dispose(bool disposing)
    {
      try
      {
        if (disposing && this.components != null)
          this.components.Dispose();
        base.Dispose(disposing);
      }
      catch (Exception ex)
      {
      }
    }

    private void InitializeComponent()
    {
            this.components = new System.ComponentModel.Container();
            this.ConsultTabControl = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.dgvData = new System.Windows.Forms.DataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.buttonDevTool = new System.Windows.Forms.Button();
            this.txtLogs = new System.Windows.Forms.RichTextBox();
            this.btnSearch = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.panelBrowser = new System.Windows.Forms.Panel();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txtUserPassword = new System.Windows.Forms.TextBox();
            this.txtUserId = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.panelReportsDownload = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.listBoxReportsDownload = new System.Windows.Forms.ListBox();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.panelReport2 = new System.Windows.Forms.Panel();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.panelReport3 = new System.Windows.Forms.Panel();
            this.tabPageQSReportCA = new System.Windows.Forms.TabPage();
            this.panelQSReportCA = new System.Windows.Forms.Panel();
            this.tabPageQSReportNC = new System.Windows.Forms.TabPage();
            this.panelQSReportNC = new System.Windows.Forms.Panel();
            this.tabPageRexReport = new System.Windows.Forms.TabPage();
            this.panelRex = new System.Windows.Forms.Panel();
            this.tabPageNewMSReportNC = new System.Windows.Forms.TabPage();
            this.panelReportNewNC = new System.Windows.Forms.Panel();
            this.tabPageNewMSReportCA = new System.Windows.Forms.TabPage();
            this.panelReportNewCA = new System.Windows.Forms.Panel();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.ConsultTabControl.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvData)).BeginInit();
            this.panel1.SuspendLayout();
            this.panelBrowser.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.panel2.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.tabPage5.SuspendLayout();
            this.tabPageQSReportCA.SuspendLayout();
            this.tabPageQSReportNC.SuspendLayout();
            this.tabPageRexReport.SuspendLayout();
            this.tabPageNewMSReportNC.SuspendLayout();
            this.tabPageNewMSReportCA.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // ConsultTabControl
            // 
            this.ConsultTabControl.Controls.Add(this.tabPage1);
            this.ConsultTabControl.Controls.Add(this.tabPage2);
            this.ConsultTabControl.Controls.Add(this.tabPage3);
            this.ConsultTabControl.Controls.Add(this.tabPage4);
            this.ConsultTabControl.Controls.Add(this.tabPage5);
            this.ConsultTabControl.Controls.Add(this.tabPageQSReportCA);
            this.ConsultTabControl.Controls.Add(this.tabPageQSReportNC);
            this.ConsultTabControl.Controls.Add(this.tabPageRexReport);
            this.ConsultTabControl.Controls.Add(this.tabPageNewMSReportNC);
            this.ConsultTabControl.Controls.Add(this.tabPageNewMSReportCA);
            this.ConsultTabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ConsultTabControl.Location = new System.Drawing.Point(0, 0);
            this.ConsultTabControl.Name = "ConsultTabControl";
            this.ConsultTabControl.SelectedIndex = 0;
            this.ConsultTabControl.Size = new System.Drawing.Size(806, 487);
            this.ConsultTabControl.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.tableLayoutPanel1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(798, 461);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Search";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 450F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.groupBox3, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.panelBrowser, 1, 0);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(3, 3);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(792, 455);
            this.tableLayoutPanel1.TabIndex = 3;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.dgvData);
            this.groupBox3.Controls.Add(this.panel1);
            this.groupBox3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox3.Location = new System.Drawing.Point(3, 3);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(444, 449);
            this.groupBox3.TabIndex = 2;
            this.groupBox3.TabStop = false;
            // 
            // dgvData
            // 
            this.dgvData.AllowUserToAddRows = false;
            this.dgvData.AllowUserToDeleteRows = false;
            this.dgvData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvData.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvData.Location = new System.Drawing.Point(3, 193);
            this.dgvData.Name = "dgvData";
            this.dgvData.ReadOnly = true;
            this.dgvData.Size = new System.Drawing.Size(438, 253);
            this.dgvData.TabIndex = 0;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.buttonDevTool);
            this.panel1.Controls.Add(this.txtLogs);
            this.panel1.Controls.Add(this.btnSearch);
            this.panel1.Controls.Add(this.progressBar1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(3, 16);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(438, 177);
            this.panel1.TabIndex = 6;
            // 
            // buttonDevTool
            // 
            this.buttonDevTool.Location = new System.Drawing.Point(364, 25);
            this.buttonDevTool.Name = "buttonDevTool";
            this.buttonDevTool.Size = new System.Drawing.Size(68, 23);
            this.buttonDevTool.TabIndex = 5;
            this.buttonDevTool.Text = "Cancel";
            this.buttonDevTool.UseVisualStyleBackColor = true;
            this.buttonDevTool.Click += new System.EventHandler(this.buttonDevTool_Click);
            // 
            // txtLogs
            // 
            this.txtLogs.Location = new System.Drawing.Point(18, 64);
            this.txtLogs.Name = "txtLogs";
            this.txtLogs.Size = new System.Drawing.Size(401, 96);
            this.txtLogs.TabIndex = 5;
            this.txtLogs.Text = "";
            // 
            // btnSearch
            // 
            this.btnSearch.Location = new System.Drawing.Point(230, 25);
            this.btnSearch.Name = "btnSearch";
            this.btnSearch.Size = new System.Drawing.Size(131, 23);
            this.btnSearch.TabIndex = 3;
            this.btnSearch.Text = "Start";
            this.btnSearch.UseVisualStyleBackColor = true;
            this.btnSearch.Click += new System.EventHandler(this.btnSearch_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(18, 25);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(214, 23);
            this.progressBar1.TabIndex = 4;
            // 
            // panelBrowser
            // 
            this.panelBrowser.Controls.Add(this.webBrowser1);
            this.panelBrowser.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelBrowser.Location = new System.Drawing.Point(453, 3);
            this.panelBrowser.Name = "panelBrowser";
            this.panelBrowser.Size = new System.Drawing.Size(336, 449);
            this.panelBrowser.TabIndex = 3;
            // 
            // webBrowser1
            // 
            this.webBrowser1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowser1.Location = new System.Drawing.Point(0, 0);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.ScriptErrorsSuppressed = true;
            this.webBrowser1.Size = new System.Drawing.Size(336, 449);
            this.webBrowser1.TabIndex = 4;
            this.webBrowser1.Url = new System.Uri("", System.UriKind.Relative);
            this.webBrowser1.Visible = false;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.groupBox1);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(798, 461);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Settings";
            this.tabPage2.UseVisualStyleBackColor = true;
            this.tabPage2.Click += new System.EventHandler(this.tabPage2_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtUserPassword);
            this.groupBox1.Controls.Add(this.txtUserId);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(25, 17);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(755, 100);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Credentials";
            // 
            // txtUserPassword
            // 
            this.txtUserPassword.Location = new System.Drawing.Point(84, 62);
            this.txtUserPassword.Name = "txtUserPassword";
            this.txtUserPassword.Size = new System.Drawing.Size(641, 20);
            this.txtUserPassword.TabIndex = 3;
            // 
            // txtUserId
            // 
            this.txtUserId.Location = new System.Drawing.Point(84, 29);
            this.txtUserId.Name = "txtUserId";
            this.txtUserId.Size = new System.Drawing.Size(641, 20);
            this.txtUserId.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(20, 65);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Password";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 32);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(40, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "UserID";
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.panelReportsDownload);
            this.tabPage3.Controls.Add(this.panel2);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(798, 461);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Report 1 Download";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // panelReportsDownload
            // 
            this.panelReportsDownload.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelReportsDownload.Location = new System.Drawing.Point(239, 0);
            this.panelReportsDownload.Name = "panelReportsDownload";
            this.panelReportsDownload.Size = new System.Drawing.Size(559, 461);
            this.panelReportsDownload.TabIndex = 1;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.listBoxReportsDownload);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(239, 461);
            this.panel2.TabIndex = 0;
            // 
            // listBoxReportsDownload
            // 
            this.listBoxReportsDownload.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listBoxReportsDownload.FormattingEnabled = true;
            this.listBoxReportsDownload.Location = new System.Drawing.Point(0, 0);
            this.listBoxReportsDownload.Name = "listBoxReportsDownload";
            this.listBoxReportsDownload.Size = new System.Drawing.Size(239, 461);
            this.listBoxReportsDownload.TabIndex = 0;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.panelReport2);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(798, 461);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Report 2 Download";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // panelReport2
            // 
            this.panelReport2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelReport2.Location = new System.Drawing.Point(0, 0);
            this.panelReport2.Name = "panelReport2";
            this.panelReport2.Size = new System.Drawing.Size(798, 461);
            this.panelReport2.TabIndex = 0;
            // 
            // tabPage5
            // 
            this.tabPage5.Controls.Add(this.panelReport3);
            this.tabPage5.Location = new System.Drawing.Point(4, 22);
            this.tabPage5.Name = "tabPage5";
            this.tabPage5.Size = new System.Drawing.Size(798, 461);
            this.tabPage5.TabIndex = 4;
            this.tabPage5.Text = "Report 3 Download";
            this.tabPage5.UseVisualStyleBackColor = true;
            // 
            // panelReport3
            // 
            this.panelReport3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelReport3.Location = new System.Drawing.Point(0, 0);
            this.panelReport3.Name = "panelReport3";
            this.panelReport3.Size = new System.Drawing.Size(798, 461);
            this.panelReport3.TabIndex = 0;
            // 
            // tabPageQSReportCA
            // 
            this.tabPageQSReportCA.Controls.Add(this.panelQSReportCA);
            this.tabPageQSReportCA.Location = new System.Drawing.Point(4, 22);
            this.tabPageQSReportCA.Name = "tabPageQSReportCA";
            this.tabPageQSReportCA.Size = new System.Drawing.Size(798, 461);
            this.tabPageQSReportCA.TabIndex = 5;
            this.tabPageQSReportCA.Text = "QS Report Stomes CA";
            this.tabPageQSReportCA.UseVisualStyleBackColor = true;
            // 
            // panelQSReportCA
            // 
            this.panelQSReportCA.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelQSReportCA.Location = new System.Drawing.Point(0, 0);
            this.panelQSReportCA.Name = "panelQSReportCA";
            this.panelQSReportCA.Size = new System.Drawing.Size(798, 461);
            this.panelQSReportCA.TabIndex = 0;
            // 
            // tabPageQSReportNC
            // 
            this.tabPageQSReportNC.Controls.Add(this.panelQSReportNC);
            this.tabPageQSReportNC.Location = new System.Drawing.Point(4, 22);
            this.tabPageQSReportNC.Name = "tabPageQSReportNC";
            this.tabPageQSReportNC.Size = new System.Drawing.Size(798, 461);
            this.tabPageQSReportNC.TabIndex = 6;
            this.tabPageQSReportNC.Text = "QS Report Stomes NC";
            this.tabPageQSReportNC.UseVisualStyleBackColor = true;
            // 
            // panelQSReportNC
            // 
            this.panelQSReportNC.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelQSReportNC.Location = new System.Drawing.Point(0, 0);
            this.panelQSReportNC.Name = "panelQSReportNC";
            this.panelQSReportNC.Size = new System.Drawing.Size(798, 461);
            this.panelQSReportNC.TabIndex = 1;
            // 
            // tabPageRexReport
            // 
            this.tabPageRexReport.Controls.Add(this.panelRex);
            this.tabPageRexReport.Location = new System.Drawing.Point(4, 22);
            this.tabPageRexReport.Name = "tabPageRexReport";
            this.tabPageRexReport.Size = new System.Drawing.Size(798, 461);
            this.tabPageRexReport.TabIndex = 7;
            this.tabPageRexReport.Text = "Rex Report";
            this.tabPageRexReport.UseVisualStyleBackColor = true;
            // 
            // panelRex
            // 
            this.panelRex.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelRex.Location = new System.Drawing.Point(0, 0);
            this.panelRex.Name = "panelRex";
            this.panelRex.Size = new System.Drawing.Size(798, 461);
            this.panelRex.TabIndex = 0;
            // 
            // tabPageNewMSReportNC
            // 
            this.tabPageNewMSReportNC.Controls.Add(this.panelReportNewNC);
            this.tabPageNewMSReportNC.Location = new System.Drawing.Point(4, 22);
            this.tabPageNewMSReportNC.Name = "tabPageNewMSReportNC";
            this.tabPageNewMSReportNC.Size = new System.Drawing.Size(798, 461);
            this.tabPageNewMSReportNC.TabIndex = 8;
            this.tabPageNewMSReportNC.Text = "New MS Report NC";
            this.tabPageNewMSReportNC.UseVisualStyleBackColor = true;
            // 
            // panelReportNewNC
            // 
            this.panelReportNewNC.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelReportNewNC.Location = new System.Drawing.Point(0, 0);
            this.panelReportNewNC.Name = "panelReportNewNC";
            this.panelReportNewNC.Size = new System.Drawing.Size(798, 461);
            this.panelReportNewNC.TabIndex = 0;
            // 
            // tabPageNewMSReportCA
            // 
            this.tabPageNewMSReportCA.Controls.Add(this.panelReportNewCA);
            this.tabPageNewMSReportCA.Location = new System.Drawing.Point(4, 22);
            this.tabPageNewMSReportCA.Name = "tabPageNewMSReportCA";
            this.tabPageNewMSReportCA.Size = new System.Drawing.Size(798, 461);
            this.tabPageNewMSReportCA.TabIndex = 9;
            this.tabPageNewMSReportCA.Text = "New MS Report CA";
            this.tabPageNewMSReportCA.UseVisualStyleBackColor = true;
            // 
            // panelReportNewCA
            // 
            this.panelReportNewCA.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelReportNewCA.Location = new System.Drawing.Point(0, 0);
            this.panelReportNewCA.Name = "panelReportNewCA";
            this.panelReportNewCA.Size = new System.Drawing.Size(798, 461);
            this.panelReportNewCA.TabIndex = 0;
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.Text = "Agents Performance System";
            this.notifyIcon1.Visible = true;
            this.notifyIcon1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon1_MouseDoubleClick);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(806, 487);
            this.Controls.Add(this.ConsultTabControl);
            this.Name = "MainForm";
            this.Text = "Agents Performance Scraping";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Resize += new System.EventHandler(this.MainForm_Resize);
            this.ConsultTabControl.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvData)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panelBrowser.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.tabPage4.ResumeLayout(false);
            this.tabPage5.ResumeLayout(false);
            this.tabPageQSReportCA.ResumeLayout(false);
            this.tabPageQSReportNC.ResumeLayout(false);
            this.tabPageRexReport.ResumeLayout(false);
            this.tabPageNewMSReportNC.ResumeLayout(false);
            this.tabPageNewMSReportCA.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.ResumeLayout(false);

    }

        private void buttonShowDevToolMS_Click(object sender, EventArgs e)
        {
            browserReport2Dev.ShowDevTools();
        }

        private void buttonReportQSCADevTool_Click(object sender, EventArgs e)
        {
            reportQSCABrowser.ShowDevTools();
        }

        private void buttonDevToolRex_Click(object sender, EventArgs e)
        {
            Control panel = null;
            ChromiumWebBrowser browser = null;

            panel = ConsultTabControl.SelectedTab.Controls.Cast<Control>().Where(c => c is Panel && c.Name.Equals("panelRex")).FirstOrDefault();
            browser = panel?.Controls.Cast<Control>().Where(c => c is ChromiumWebBrowser).FirstOrDefault() as ChromiumWebBrowser;

            browser.ShowDevTools();
        }
    }
}
