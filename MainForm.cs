// Decompiled with JetBrains decompiler
// Type: Dax.Scrapping.Appraisal.MainForm
// Assembly: Dax.Scrapping.Appraisal, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 636504A0-34C1-4208-9BC8-AA1181CC609A
// Assembly location: C:\Users\manue\Desktop\EsubScraper\Dax.Scrapping.Appraisal.exe

using CefSharp;
using CefSharp.WinForms;
using Dax.Scrapping.Appraisal.Core;
using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Timers;
using System.Windows.Forms;

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
    private NotifyIcon notifyIcon1;

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

    private void MainForm_Load(object sender, EventArgs e)
    {
      this.LoadConfiguration();
      this.InitializeBrouser();
      Cef.Initialize();
      double num = Convert.ToDouble(Helper.GetAppSettingAsString("Time"));
      this.aTimer.Elapsed += new ElapsedEventHandler(this.ATimer_Elapsed);
      this.aTimer.Interval = num;
      this.aTimer.Enabled = true;
      this.Clear();
      this.panelBrowser.Controls.Clear();
      this.btnSearch_Click((object) null, (EventArgs) null);
    }

    private void ATimer_Elapsed(object sender, ElapsedEventArgs e)
    {
        this._scraper._curStatus = Status.Searching;
        this._scraper.LoadWeb(this._scraper._newAgentsInfoSite);
        //this._brouserComponent.Load(this._newAgentsInfoSite);
        //Application.Restart();
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
      }
      catch (Exception ex)
      {
        //int num = (int) MessageBox.Show(ex.Message);
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
      this.components = (IContainer) new Container();
      ComponentResourceManager componentResourceManager = new ComponentResourceManager(typeof (MainForm));
      this.ConsultTabControl = new TabControl();
      this.tabPage1 = new TabPage();
      this.tableLayoutPanel1 = new TableLayoutPanel();
      this.groupBox3 = new GroupBox();
      this.dgvData = new DataGridView();
      this.panel1 = new Panel();
      this.buttonDevTool = new Button();
      this.txtLogs = new RichTextBox();
      this.btnSearch = new Button();
      this.progressBar1 = new ProgressBar();
      this.panelBrowser = new Panel();
      this.webBrowser1 = new WebBrowser();
      this.tabPage2 = new TabPage();
      this.groupBox1 = new GroupBox();
      this.txtUserPassword = new TextBox();
      this.txtUserId = new TextBox();
      this.label2 = new Label();
      this.label1 = new Label();
      this.bindingSource1 = new BindingSource(this.components);
      this.notifyIcon1 = new NotifyIcon(this.components);
      this.ConsultTabControl.SuspendLayout();
      this.tabPage1.SuspendLayout();
      this.tableLayoutPanel1.SuspendLayout();
      this.groupBox3.SuspendLayout();
      ((ISupportInitialize) this.dgvData).BeginInit();
      this.panel1.SuspendLayout();
      this.panelBrowser.SuspendLayout();
      this.tabPage2.SuspendLayout();
      this.groupBox1.SuspendLayout();
      ((ISupportInitialize) this.bindingSource1).BeginInit();
      this.SuspendLayout();
      this.ConsultTabControl.Controls.Add((Control) this.tabPage1);
      this.ConsultTabControl.Controls.Add((Control) this.tabPage2);
      this.ConsultTabControl.Dock = DockStyle.Fill;
      this.ConsultTabControl.Location = new Point(0, 0);
      this.ConsultTabControl.Name = "ConsultTabControl";
      this.ConsultTabControl.SelectedIndex = 0;
      this.ConsultTabControl.Size = new Size(806, 487);
      this.ConsultTabControl.TabIndex = 0;
      this.tabPage1.Controls.Add((Control) this.tableLayoutPanel1);
      this.tabPage1.Location = new Point(4, 22);
      this.tabPage1.Name = "tabPage1";
      this.tabPage1.Padding = new Padding(3);
      this.tabPage1.Size = new Size(798, 461);
      this.tabPage1.TabIndex = 0;
      this.tabPage1.Text = "Search";
      this.tabPage1.UseVisualStyleBackColor = true;
      this.tableLayoutPanel1.ColumnCount = 2;
      this.tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 450f));
      this.tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
      this.tableLayoutPanel1.Controls.Add((Control) this.groupBox3, 0, 0);
      this.tableLayoutPanel1.Controls.Add((Control) this.panelBrowser, 1, 0);
      this.tableLayoutPanel1.Dock = DockStyle.Fill;
      this.tableLayoutPanel1.Location = new Point(3, 3);
      this.tableLayoutPanel1.Name = "tableLayoutPanel1";
      this.tableLayoutPanel1.RowCount = 1;
      this.tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
      this.tableLayoutPanel1.Size = new Size(792, 455);
      this.tableLayoutPanel1.TabIndex = 3;
      this.groupBox3.Controls.Add((Control) this.dgvData);
      this.groupBox3.Controls.Add((Control) this.panel1);
      this.groupBox3.Dock = DockStyle.Fill;
      this.groupBox3.Location = new Point(3, 3);
      this.groupBox3.Name = "groupBox3";
      this.groupBox3.Size = new Size(444, 449);
      this.groupBox3.TabIndex = 2;
      this.groupBox3.TabStop = false;
      this.dgvData.AllowUserToAddRows = false;
      this.dgvData.AllowUserToDeleteRows = false;
      this.dgvData.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
      this.dgvData.Dock = DockStyle.Fill;
      this.dgvData.Location = new Point(3, 193);
      this.dgvData.Name = "dgvData";
      this.dgvData.ReadOnly = true;
      this.dgvData.Size = new Size(438, 253);
      this.dgvData.TabIndex = 0;
      this.panel1.Controls.Add((Control) this.buttonDevTool);
      this.panel1.Controls.Add((Control) this.txtLogs);
      this.panel1.Controls.Add((Control) this.btnSearch);
      this.panel1.Controls.Add((Control) this.progressBar1);
      this.panel1.Dock = DockStyle.Top;
      this.panel1.Location = new Point(3, 16);
      this.panel1.Name = "panel1";
      this.panel1.Size = new Size(438, 177);
      this.panel1.TabIndex = 6;
      this.buttonDevTool.Location = new Point(364, 25);
      this.buttonDevTool.Name = "buttonDevTool";
      this.buttonDevTool.Size = new Size(68, 23);
      this.buttonDevTool.TabIndex = 5;
      this.buttonDevTool.Text = "Cancel";
      this.buttonDevTool.UseVisualStyleBackColor = true;
      this.buttonDevTool.Click += new EventHandler(this.buttonDevTool_Click);
      this.txtLogs.Location = new Point(18, 64);
      this.txtLogs.Name = "txtLogs";
      this.txtLogs.Size = new Size(401, 96);
      this.txtLogs.TabIndex = 5;
      this.txtLogs.Text = "";
      this.btnSearch.Location = new Point(230, 25);
      this.btnSearch.Name = "btnSearch";
      this.btnSearch.Size = new Size(131, 23);
      this.btnSearch.TabIndex = 3;
      this.btnSearch.Text = "Start";
      this.btnSearch.UseVisualStyleBackColor = true;
      this.btnSearch.Click += new EventHandler(this.btnSearch_Click);
      this.progressBar1.Location = new Point(18, 25);
      this.progressBar1.Name = "progressBar1";
      this.progressBar1.Size = new Size(214, 23);
      this.progressBar1.TabIndex = 4;
      this.panelBrowser.Controls.Add((Control) this.webBrowser1);
      this.panelBrowser.Dock = DockStyle.Fill;
      this.panelBrowser.Location = new Point(453, 3);
      this.panelBrowser.Name = "panelBrowser";
      this.panelBrowser.Size = new Size(336, 449);
      this.panelBrowser.TabIndex = 3;
      this.webBrowser1.Dock = DockStyle.Fill;
      this.webBrowser1.Location = new Point(0, 0);
      this.webBrowser1.MinimumSize = new Size(20, 20);
      this.webBrowser1.Name = "webBrowser1";
      this.webBrowser1.ScriptErrorsSuppressed = true;
      this.webBrowser1.Size = new Size(336, 449);
      this.webBrowser1.TabIndex = 4;
      this.webBrowser1.Url = new Uri("", UriKind.Relative);
      this.webBrowser1.Visible = false;
      this.tabPage2.Controls.Add((Control) this.groupBox1);
      this.tabPage2.Location = new Point(4, 22);
      this.tabPage2.Name = "tabPage2";
      this.tabPage2.Padding = new Padding(3);
      this.tabPage2.Size = new Size(798, 461);
      this.tabPage2.TabIndex = 1;
      this.tabPage2.Text = "Settings";
      this.tabPage2.UseVisualStyleBackColor = true;
      this.tabPage2.Click += new EventHandler(this.tabPage2_Click);
      this.groupBox1.Controls.Add((Control) this.txtUserPassword);
      this.groupBox1.Controls.Add((Control) this.txtUserId);
      this.groupBox1.Controls.Add((Control) this.label2);
      this.groupBox1.Controls.Add((Control) this.label1);
      this.groupBox1.Location = new Point(25, 17);
      this.groupBox1.Name = "groupBox1";
      this.groupBox1.Size = new Size(755, 100);
      this.groupBox1.TabIndex = 0;
      this.groupBox1.TabStop = false;
      this.groupBox1.Text = "Credentials";
      this.txtUserPassword.Location = new Point(84, 62);
      this.txtUserPassword.Name = "txtUserPassword";
      this.txtUserPassword.Size = new Size(641, 20);
      this.txtUserPassword.TabIndex = 3;
      this.txtUserId.Location = new Point(84, 29);
      this.txtUserId.Name = "txtUserId";
      this.txtUserId.Size = new Size(641, 20);
      this.txtUserId.TabIndex = 2;
      this.label2.AutoSize = true;
      this.label2.Location = new Point(20, 65);
      this.label2.Name = "label2";
      this.label2.Size = new Size(53, 13);
      this.label2.TabIndex = 1;
      this.label2.Text = "Password";
      this.label1.AutoSize = true;
      this.label1.Location = new Point(20, 32);
      this.label1.Name = "label1";
      this.label1.Size = new Size(40, 13);
      this.label1.TabIndex = 0;
      this.label1.Text = "UserID";
      this.notifyIcon1.Icon = (Icon) componentResourceManager.GetObject("notifyIcon1.Icon");
      this.notifyIcon1.Text = "Agents Performance System";
      this.notifyIcon1.Visible = true;
      this.notifyIcon1.MouseDoubleClick += new MouseEventHandler(this.notifyIcon1_MouseDoubleClick);
      this.AutoScaleDimensions = new SizeF(6f, 13f);
      this.AutoScaleMode = AutoScaleMode.Font;
      this.ClientSize = new Size(806, 487);
      this.Controls.Add((Control) this.ConsultTabControl);
      this.Icon = (Icon) componentResourceManager.GetObject("$this.Icon");
      this.Name = "MainForm";
      this.Text = "Agents Performance Scraping";
      this.WindowState = FormWindowState.Maximized;
      this.FormClosing += new FormClosingEventHandler(this.MainForm_FormClosing);
      this.Load += new EventHandler(this.MainForm_Load);
      this.Resize += new EventHandler(this.MainForm_Resize);
      this.ConsultTabControl.ResumeLayout(false);
      this.tabPage1.ResumeLayout(false);
      this.tableLayoutPanel1.ResumeLayout(false);
      this.groupBox3.ResumeLayout(false);
      ((ISupportInitialize) this.dgvData).EndInit();
      this.panel1.ResumeLayout(false);
      this.panelBrowser.ResumeLayout(false);
      this.tabPage2.ResumeLayout(false);
      this.groupBox1.ResumeLayout(false);
      this.groupBox1.PerformLayout();
      ((ISupportInitialize) this.bindingSource1).EndInit();
      this.ResumeLayout(false);
    }
  }
}
