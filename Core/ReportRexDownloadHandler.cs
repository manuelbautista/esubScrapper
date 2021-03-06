﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CefSharp;

namespace Dax.Scrapping.Appraisal.Core
{
    public class ReportRexDownloadHandler : IDownloadHandler
    {
        public string filePath
        {
            get
            {
                var filePath = @"C:\esubmitter\Reports\ReportRex\";
                if (!Directory.Exists(filePath))
                {
                    Directory.CreateDirectory(filePath);
                }
                return filePath;
            }
        }

        public string _fileName { get; set; }

        private bool FileExists(string fileName)
        {
            var school = new SchoolEntities();
            return school.DailyEmailReports.ToList().Any(a => a.ReportName.Equals(fileName) && a.Date.HasValue && a.Date.Value.Equals(DateTime.Now.Date));
        }
        public void OnBeforeDownload(IBrowser browser, DownloadItem downloadItem, IBeforeDownloadCallback callback)
        {

            if (!callback.IsDisposed)
            {
                var yesterday = DateTime.Now.Date.AddDays(-1);
                string date = yesterday.ToString("yyyy-MM-dd");
                string fileName = "Rex_StomesNC_";

                downloadItem.SuggestedFileName = fileName + date + ".csv";
                _fileName = downloadItem.SuggestedFileName;

                using (callback)
                {
                    //
                    callback.Continue(filePath + downloadItem.SuggestedFileName, showDialog: false);
                }
            }
        }


        public void OnDownloadUpdated(IBrowser browser, DownloadItem downloadItem, IDownloadItemCallback callback)
        {
            if (downloadItem.IsComplete)
            {
                //Remove previous report
                Helper.RemoveReport("ReportRex");

                var school = new SchoolEntities();
                var daily = new DailyEmailReport
                {
                    Time = DateTime.Now.ToShortTimeString(),
                    Date = DateTime.Now.Date,
                    Path = filePath + _fileName,
                    ReportName = "ReportRex",
                    Sent = false
                };
                school.DailyEmailReports.Add(daily);
                school.SaveChanges();
            }
        }
    }
}
