//------------------------------------------------------------------------------
// <auto-generated>
//    This code was generated from a template.
//
//    Manual changes to this file may cause unexpected behavior in your application.
//    Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace InventoryFeedService
{
    using System;
    using System.Collections.Generic;
    
    public partial class tblInventoryFeedProcess
    {
        public int ifp_id { get; set; }
        public Nullable<int> if_id { get; set; }
        public Nullable<System.TimeSpan> time_split { get; set; }
        public Nullable<System.DateTime> datetime_updated { get; set; }
        public string status { get; set; }
        public Nullable<int> current_pr { get; set; }
    }
}
