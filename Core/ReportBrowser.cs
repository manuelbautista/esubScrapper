using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CefSharp.WinForms;

namespace Dax.Scrapping.Appraisal.Core
{
    public class ReportBrowser:ChromiumWebBrowser
    {
        //protected ChromiumWebBrowser _browserComponent;

        //public ChromiumWebBrowser BrowserComponent
        //{
        //    get
        //    {
        //        return this._browserComponent;
        //    }
        //}

        public ReportBrowser(string address) : base(address)
        {
            this.Load(address);
        }
    }
}
