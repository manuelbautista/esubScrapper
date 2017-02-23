﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CefSharp;

namespace Dax.Scrapping.Appraisal.Core
{
    public class ReportNewNCDownloadHandler : IDownloadHandler
    {
        private bool FileExists(string fileName)
        {
            var school = new SchoolEntities();
            return school.DailyEmailReports.ToList().Any(a => a.ReportName.Equals(fileName) && a.Date.HasValue && a.Date.Value.Equals(DateTime.Now.Date));
        }
        public void OnBeforeDownload(IBrowser browser, DownloadItem downloadItem, IBeforeDownloadCallback callback)
        {
            var filePath = @"C:\esubmitter\Reports\ReportNewNC\";
            if (!Directory.Exists(filePath))
            {
                Directory.CreateDirectory(filePath);
            }
            if (!callback.IsDisposed)
            {
                downloadItem.SuggestedFileName = "(NC)" + downloadItem.SuggestedFileName;

                //if (FileExists(downloadItem.SuggestedFileName)) return;

                using (callback)
                {
                    //Remove previous report
                    Helper.RemoveReport("ReportNewNC");
                    //
                    callback.Continue(filePath + downloadItem.SuggestedFileName, showDialog: false);

                    var school = new SchoolEntities();
                    var daily = new DailyEmailReport
                    {
                        Time = DateTime.Now.ToShortTimeString(),
                        Date = DateTime.Now.Date,
                        Path = filePath + downloadItem.SuggestedFileName,
                        ReportName = "ReportNewNC",
                        Sent = false
                    };
                    school.DailyEmailReports.Add(daily);
                    school.SaveChanges();
                }
            }
        }


        public void OnDownloadUpdated(IBrowser browser, DownloadItem downloadItem, IDownloadItemCallback callback)
        {
            //throw new NotImplementedException();
        }
    }
}
