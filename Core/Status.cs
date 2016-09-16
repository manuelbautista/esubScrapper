// Decompiled with JetBrains decompiler
// Type: Dax.Scrapping.Appraisal.Core.Status
// Assembly: Dax.Scrapping.Appraisal, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 636504A0-34C1-4208-9BC8-AA1181CC609A
// Assembly location: C:\Users\manue\Desktop\EsubScraper\Dax.Scrapping.Appraisal.exe

using System.ComponentModel;

namespace Dax.Scrapping.Appraisal.Core
{
  public enum Status
  {
    [Description("Paused")] Paused,
    [Description("loading login screen")] Loading,
    [Description("Loging user")] Loggin,
    [Description("Searching Conections")] Searching,
    [Description("Parsing Conections page")] Parsing,
    [Description("Getting Orders Info")] GetAgentsInfo,
    [Description("Completed")] Completed,
  }
}
