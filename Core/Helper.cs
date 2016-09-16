// Decompiled with JetBrains decompiler
// Type: Dax.Scrapping.Appraisal.Core.Helper
// Assembly: Dax.Scrapping.Appraisal, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 636504A0-34C1-4208-9BC8-AA1181CC609A
// Assembly location: C:\Users\manue\Desktop\EsubScraper\Dax.Scrapping.Appraisal.exe

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Dax.Scrapping.Appraisal.Core
{
  public static class Helper
  {
    private static AppSettingsReader reader = new AppSettingsReader();
    private static string _serverHub = (string) null;

    public static string ServerHub
    {
      get
      {
        return Helper._serverHub;
      }
      set
      {
        Helper._serverHub = value;
      }
    }

    public static string GetAppSettingAsString(string key)
    {
      return ConfigurationManager.AppSettings.Get(key);
    }

    public static int GetAppSettingAsInt(string key)
    {
      return Convert.ToInt32(ConfigurationManager.AppSettings.Get(key));
    }

    public static string GetServerUrl()
    {
      string str = Helper._serverHub ?? Helper.GetAppSettingAsString("ServerUrl").Trim();
      if (!str.EndsWith("/"))
        str += "/";
      return str;
    }

    public static void Modify(string key, string value)
    {
      System.Configuration.Configuration configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
      configuration.AppSettings.Settings[key].Value = value;
      configuration.Save();
      ConfigurationManager.RefreshSection("appSettings");
    }

    public static string GetEnumDescription<T>(T value)
    {
      DescriptionAttribute[] customAttributes = (DescriptionAttribute[]) value.GetType().GetField(value.ToString()).GetCustomAttributes(typeof (DescriptionAttribute), false);
      if (customAttributes != null && (uint) customAttributes.Length > 0U)
        return customAttributes[0].Description;
      return value.ToString();
    }

    public static string GetFileContent(string path)
    {
      if (!File.Exists(path))
        throw new FileNotFoundException();
      return File.ReadAllText(path);
    }

    public static void WriteCSV<T>(IEnumerable<T> items, string path)
    {
      PropertyInfo[] properties = typeof (T).GetProperties(BindingFlags.Instance | BindingFlags.Public);
      using (StreamWriter streamWriter = new StreamWriter(path))
      {
        streamWriter.WriteLine(string.Join("| ", ((IEnumerable<PropertyInfo>) properties).Select<PropertyInfo, string>((Func<PropertyInfo, string>) (p => p.Name))));
        foreach (T obj in items)
        {
          T item = obj;
          streamWriter.WriteLine(string.Join<object>("| ", ((IEnumerable<PropertyInfo>) properties).Select<PropertyInfo, object>((Func<PropertyInfo, object>) (p => p.GetValue((object) item, (object[]) null)))));
        }
      }
    }
  }
}
