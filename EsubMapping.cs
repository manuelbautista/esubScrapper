//------------------------------------------------------------------------------
// <auto-generated>
//    This code was generated from a template.
//
//    Manual changes to this file may cause unexpected behavior in your application.
//    Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace Dax.Scrapping.Appraisal
{
    using System;
    using System.Collections.Generic;
    
    public partial class EsubMapping
    {
        public EsubMapping()
        {
            this.EsubMappingValues = new HashSet<EsubMappingValue>();
        }
    
        public string AreaofStudy { get; set; }
        public string esubListing { get; set; }
        public string CCMapping { get; set; }
        public string MSEMapping { get; set; }
        public string MCM1Mapping { get; set; }
        public string MCM2Mapping { get; set; }
        public string MCM3Mapping { get; set; }
        public string MSEMajorMapping { get; set; }
        public int Id { get; set; }
    
        public virtual ICollection<EsubMappingValue> EsubMappingValues { get; set; }
    }
}
