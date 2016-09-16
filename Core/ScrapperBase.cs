// Decompiled with JetBrains decompiler
// Type: Dax.Scrapping.Appraisal.Core.ScrapperBase
// Assembly: Dax.Scrapping.Appraisal, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 636504A0-34C1-4208-9BC8-AA1181CC609A
// Assembly location: C:\Users\manue\Desktop\EsubScraper\Dax.Scrapping.Appraisal.exe

using CefSharp;
using CefSharp.WinForms;
using System.Threading;

namespace Dax.Scrapping.Appraisal.Core
{
  public class ScrapperBase
  {
    protected ChromiumWebBrowser _brouserComponent;

    public ChromiumWebBrowser BrouserComponent
    {
      get
      {
        return this._brouserComponent;
      }
    }

    protected void Click(string id)
    {
      this._brouserComponent.ExecuteScriptAsync("(function() {\n                                    document.getElementById('{0}').click();\n                               })();".Replace("{0}", id));
    }

    protected void ScrollDown()
    {
      this._brouserComponent.ExecuteScriptAsync("(function() {                                        \n                                    window.scrollTo(0, document.body.scrollHeight || document.documentElement.scrollHeight);\n                               })();");
      Thread.Sleep(2000);
    }
  }
}
