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
    
    public partial class user
    {
        public int id { get; set; }
        public string username { get; set; }
        public string password { get; set; }
        public Nullable<System.DateTime> last_login { get; set; }
        public Nullable<System.DateTime> create_date { get; set; }
        public string login_ip { get; set; }
        public string user_type { get; set; }
        public string user_group { get; set; }
        public string Fullname { get; set; }
        public string DialerID { get; set; }
        public string CCAccess { get; set; }
        public string MSAccess { get; set; }
        public string HEGAccess { get; set; }
        public string QACheck { get; set; }
        public Nullable<int> QAFrequency { get; set; }
        public Nullable<System.DateTime> ActualDate { get; set; }
        public Nullable<int> QACount { get; set; }
        public Nullable<int> QAClientId { get; set; }
        public Nullable<int> QAClientCount { get; set; }
        public Nullable<System.DateTime> FirstLogin { get; set; }
        public Nullable<System.DateTime> LastLogout { get; set; }
        public Nullable<bool> IsLoggedIn { get; set; }
        public string EsubCurrentVersion { get; set; }
        public string HEGCAccess { get; set; }
        public string MSEAccess { get; set; }
        public string MCMAcess { get; set; }
        public Nullable<bool> IsActive { get; set; }
        public Nullable<bool> Lockout { get; set; }
    }
}
