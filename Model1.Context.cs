// Decompiled with JetBrains decompiler
// Type: Dax.Scrapping.Appraisal.ScrapingEntities
// Assembly: Dax.Scrapping.Appraisal, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null
// MVID: 636504A0-34C1-4208-9BC8-AA1181CC609A
// Assembly location: C:\Users\manue\Desktop\EsubScraper\Dax.Scrapping.Appraisal.exe

using System.Data.Entity;
using System.Data.Entity.Infrastructure;

namespace Dax.Scrapping.Appraisal
{
  public class ScrapingEntities : DbContext
  {
    public DbSet<AgentsInfo> AgentsInfoes { get; set; }

    public ScrapingEntities()
      : base("name=ScrapingEntities")
    {
    }

    protected override void OnModelCreating(DbModelBuilder modelBuilder)
    {
      throw new UnintentionalCodeFirstException();
    }
  }
}
