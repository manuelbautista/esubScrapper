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
using System.Data.Entity.Migrations;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Threading;
using System.Timers;
using System.Windows.Forms;
using EAGetMail;

namespace Dax.Scrapping.Appraisal
{
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
        public StepsEnum.Steps actualStepReport3;
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
            //// Create a folder named "inbox" under current directory
            //// to save the email retrieved.
            string curpath = Directory.GetCurrentDirectory();
            List<Mail> emails = new List<Mail>();

            //string mailbox = String.Format("{0}\\inbox", curpath);

            //// If the folder is not existed, create it.
            //if (!Directory.Exists(mailbox))
            //{
            //    Directory.CreateDirectory(mailbox);
            //}

            MailServer oServer = new MailServer("pop.secureserver.net",
                        "reports@statewideconsultants.com", "Educ@t|09", ServerProtocol.Pop3);
            MailClient oClient = new MailClient("TryIt");

            // If your POP3 server requires SSL connection,
            // Please add the following codes:
            // oServer.SSLConnection = true;
            // oServer.Port = 995;

            try
            {
                oClient.Connect(oServer);
                MailInfo[] infos = oClient.GetMailInfos();
                for (int i = 0; i < infos.Length; i++)
                {
                    MailInfo info = infos[i];
                    Console.WriteLine("Index: {0}; Size: {1}; UIDL: {2}",
                        info.Index, info.Size, info.UIDL);

                    // Receive email from POP3 server
                    Mail oMail = oClient.GetMail(info);
                    emails.Add(oMail);

                }

                // Quit and pure emails marked as deleted from POP3 server.
                oClient.Quit();
                var todayEmails = emails.Where(a => a.ReceivedDate.Date.Equals(DateTime.Now.Date));
                foreach (var email in todayEmails)
                {
                    if (email.Subject.Contains("Export Call Report"))
                    {
                        var url = sep(email.HtmlBody);
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
            return "";
             
        }
        private void SendReport3ToEmail(string subject)
        {
            //send the first Report
            var daily = new SchoolEntities();
            var reportsToSend = daily.DailyEmailReports.Where(a => a.Sent.HasValue && a.Sent.Value.Equals(false)).ToList().Where(a => a.Date.HasValue && a.Date.Value.Equals(DateTime.Now.Date) && a.ReportName.Equals("Report3"));

            foreach (var report in reportsToSend)
            {
                //MailMessage mail = new MailMessage();
                MailMessage mail = new MailMessage("reports@statewideconsultants.com", "reports@statewideconsultants.com");
                SmtpClient client = new SmtpClient();
                client.Host = "smtpout.secureserver.net";

                client.Port = 3535;  //Tried 80, 3535, 25, 465 (SSL)
                client.UseDefaultCredentials = false;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.EnableSsl = false;
                client.Credentials = new NetworkCredential("reports@statewideconsultants.com", "Educ@t|09");
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
            }
        }

        private void SendReport2ToEmail(string subject)
        {
            //send the first Report
            var daily = new SchoolEntities();
            var reportsToSend = daily.DailyEmailReports.Where(a => a.Sent.HasValue && a.Sent.Value.Equals(false)).ToList().Where(a => a.Date.HasValue && a.Date.Value.Equals(DateTime.Now.Date) && a.ReportName.Equals("Report2"));

            foreach (var report in reportsToSend)
            {
                //MailMessage mail = new MailMessage();
                MailMessage mail = new MailMessage("reports@statewideconsultants.com", "reports@statewideconsultants.com");
                SmtpClient client = new SmtpClient();
                client.Host = "smtpout.secureserver.net";

                client.Port = 3535;  //Tried 80, 3535, 25, 465 (SSL)
                client.UseDefaultCredentials = false;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.EnableSsl = false;
                client.Credentials = new NetworkCredential("reports@statewideconsultants.com", "Educ@t|09");
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
            }
        }

        private void SendReport1ToEmail(string subject)
    {
        //send the first Report
        var daily = new SchoolEntities();
        var reportsToSend = daily.DailyEmailReports.Where(a => a.Sent.HasValue && a.Sent.Value.Equals(false)).ToList().Where(a=> a.Date.HasValue && a.Date.Value.Equals(DateTime.Now.Date) && a.ReportName.Equals("Report1"));

        foreach (var report in reportsToSend)
        {
            //MailMessage mail = new MailMessage();
            MailMessage mail = new MailMessage("reports@statewideconsultants.com", "reports@statewideconsultants.com");
            SmtpClient client = new SmtpClient();
            client.Host = "smtpout.secureserver.net";

            client.Port = 3535;  //Tried 80, 3535, 25, 465 (SSL)
            client.UseDefaultCredentials = false;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.EnableSsl = false;
            client.Credentials = new NetworkCredential("reports@statewideconsultants.com", "Educ@t|09");
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
        }
    }

    private void MainForm_Load(object sender, EventArgs e)
    {
      this.LoadConfiguration();
      this.InitializeBrouser();
        var settings = new CefSettings();
        //Property to fix the custom scaling problem
        //settings.CefCommandLineArgs.Add("proxy-server", "107.182.113.247:56227");
        //settings.CefCommandLineArgs.Add("proxy-bypass-list", "127.*,192.168.*,10.10.*,193.9.162.*");
        //Cef.Initialize(settings);
        Cef.Initialize();

        double num = Convert.ToDouble(Helper.GetAppSettingAsString("Time"));
        this.aTimer.Elapsed += new ElapsedEventHandler(this.ATimer_Elapsed);
        this.aTimer.Interval = num;
        this.aTimer.Enabled = true;
        this.Clear();
        this.panelBrowser.Controls.Clear();
        this.btnSearch_Click((object) null, (EventArgs) null);
    }

    private bool Report1IsSent()
    {
        bool isSent = true;
        if (DateTime.Now.Hour >= 9 && DateTime.Now.Hour < 17  && DateTime.Now.DayOfWeek != DayOfWeek.Sunday)
        {
            SchoolEntities school = new SchoolEntities();
             isSent =
                school.DailyEmailReports.ToList()
                    .Any(
                        a =>
                            a.Date.Equals(DateTime.Now.Date) && a.ReportName.Equals("Report1") && a.Sent.HasValue &&
                            a.Sent.Value.Equals(true));
        }
        return isSent;
    }
    private bool SaveReport1Today()
    {
        SchoolEntities school = new SchoolEntities();
        var isSent = school.DailyEmailReports.ToList().Any(a => a.Date.Equals(DateTime.Now.Date) && a.ReportName.Equals("Report1"));

        return !isSent;
    }
    private bool Report2IsSent()
    {
        bool isSent = true;
        if (DateTime.Now.Hour >= 9 && DateTime.Now.Hour < 17  && DateTime.Now.DayOfWeek != DayOfWeek.Sunday)
        {
            SchoolEntities school = new SchoolEntities();
            isSent =
                school.DailyEmailReports.ToList()
                    .Any(
                        a =>
                            a.Date.Equals(DateTime.Now.Date) && a.ReportName.Equals("Report2") && a.Sent.HasValue &&
                            a.Sent.Value.Equals(true));
        }
        return isSent;
    }
    private bool SaveReport2Today()
        {
            SchoolEntities school = new SchoolEntities();
            var isSent = school.DailyEmailReports.ToList().Any(a => a.Date.Equals(DateTime.Now.Date) && a.ReportName.Equals("Report2"));

            return !isSent;
        }
    private bool Report3IsSent()
    {
        bool isSent = true;
        if (DateTime.Now.Hour >= 9 && DateTime.Now.Hour < 17  && DateTime.Now.DayOfWeek != DayOfWeek.Sunday)
        {
            SchoolEntities school = new SchoolEntities();
            isSent =
                school.DailyEmailReports.ToList()
                    .Any(
                        a =>
                            a.Date.Equals(DateTime.Now.Date) && a.ReportName.Equals("Report3") && a.Sent.HasValue &&
                            a.Sent.Value.Equals(true));
        }
        return isSent;
    }
    private bool SaveReport3Today()
        {
            SchoolEntities school = new SchoolEntities();
            var isSent = school.DailyEmailReports.ToList().Any(a => a.Date.Equals(DateTime.Now.Date) && a.ReportName.Equals("Report3"));

            return !isSent;
        }


      private void ATimer_Elapsed(object sender, ElapsedEventArgs e)
      {
          this._scraper._curStatus = Status.Searching;
          this._scraper.LoadWeb(this._scraper._newAgentsInfoSite);

          actualStepReport2 = StepsEnum.Steps.Step1;
          actualStepReport3 = StepsEnum.Steps.Step1;

          if (DateTime.Now.Hour == 9)
          {
              var scrapper = new ScrapperRestart();
              var school = new SchoolEntities();
              if (!school.ScrapperRestarts.ToList().Any(a => a.Date.HasValue && a.Date.Value.Equals(DateTime.Now.Date)))
              {
                  scrapper.Date = DateTime.Now.Date;
                  scrapper.Restarted = true;
                  school.ScrapperRestarts.Add(scrapper);
                  school.SaveChanges();
                   //Restart the application
                  Application.Restart();
              }
          }
          //just execute the download and send file between 9 and 4
          if (DateTime.Now.Hour >= 9 && DateTime.Now.Hour < 17  && DateTime.Now.DayOfWeek != DayOfWeek.Sunday)
          {
                //Send reports info
                SaveReports();
                //Send Emails
                SendReport1ToEmail("Export Call Report");
                SendReport2ToEmail("MS CA Report");
                SendReport3ToEmail("MS NC Report");
            }
      }

      private bool recordReport3IsSaved()
        {
            SchoolEntities school = new SchoolEntities();
            return
                school.DailyEmailReports
                    .ToList()
                    .Any(a => a.Date.HasValue && a.Date.Value.Equals(DateTime.Now.Date) && a.ReportName.Equals("Report3"));
        }
        private bool recordReport2IsSaved()
    {
        SchoolEntities school = new SchoolEntities();
        return
            school.DailyEmailReports
                .ToList()
                .Any(a => a.Date.HasValue && a.Date.Value.Equals(DateTime.Now.Date) && a.ReportName.Equals("Report2"));
    }
    private bool recordIsSaved()
    {
        SchoolEntities school = new SchoolEntities();
        return
            school.DailyEmailReports
                .ToList()
                .Any(a => a.Date.HasValue && a.Date.Value.Equals(DateTime.Now.Date) && a.ReportName.Equals("Report1"));
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
                SaveReports();
      }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
    }

      private void SaveReports()
      {
            bool saveReport = SaveReport1Today();
            //bool reportIsSent = Report1IsSent();
            //
            bool saveReport2 = SaveReport2Today();
            //bool report2IsSent = Report2IsSent();
            //
            bool saveReport3 = SaveReport3Today();
            //bool report3IsSent = Report3IsSent();

            if (saveReport)
            {
                var urlReport = GetReport1LinkFromEmail();
                if (!string.IsNullOrEmpty(urlReport))
                {
                    ChromiumWebBrowser browserReport1 = new ReportBrowser(urlReport);
                    browserReport1.DownloadHandler = new ReportDownloadHandler();
                    this.panelReportsDownload.Controls.Clear();
                    this.panelReportsDownload.Controls.Add(browserReport1);
                    browserReport1.LoadingStateChanged += BrowserReport1_LoadingStateChanged;
                    ConsultTabControl.SelectedIndex = 2;
                    Thread.Sleep(1000);
                    ConsultTabControl.SelectedIndex = 0;
                }
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
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panelReportsDownload = new System.Windows.Forms.Panel();
            this.listBoxReportsDownload = new System.Windows.Forms.ListBox();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.panelReport2 = new System.Windows.Forms.Panel();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.panelReport3 = new System.Windows.Forms.Panel();
            this.ConsultTabControl.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvData)).BeginInit();
            this.panel1.SuspendLayout();
            this.panelBrowser.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.tabPage3.SuspendLayout();
            this.panel2.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.tabPage5.SuspendLayout();
            this.SuspendLayout();
            // 
            // ConsultTabControl
            // 
            this.ConsultTabControl.Controls.Add(this.tabPage1);
            this.ConsultTabControl.Controls.Add(this.tabPage2);
            this.ConsultTabControl.Controls.Add(this.tabPage3);
            this.ConsultTabControl.Controls.Add(this.tabPage4);
            this.ConsultTabControl.Controls.Add(this.tabPage5);
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
            // notifyIcon1
            // 
            this.notifyIcon1.Text = "Agents Performance System";
            this.notifyIcon1.Visible = true;
            this.notifyIcon1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.notifyIcon1_MouseDoubleClick);
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
            // panel2
            // 
            this.panel2.Controls.Add(this.listBoxReportsDownload);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(239, 461);
            this.panel2.TabIndex = 0;
            // 
            // panelReportsDownload
            // 
            this.panelReportsDownload.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelReportsDownload.Location = new System.Drawing.Point(239, 0);
            this.panelReportsDownload.Name = "panelReportsDownload";
            this.panelReportsDownload.Size = new System.Drawing.Size(559, 461);
            this.panelReportsDownload.TabIndex = 1;
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
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.tabPage3.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.tabPage4.ResumeLayout(false);
            this.tabPage5.ResumeLayout(false);
            this.ResumeLayout(false);

    }

        private void buttonShowDevToolMS_Click(object sender, EventArgs e)
        {
            browserReport2Dev.ShowDevTools();
        }
    }
}
