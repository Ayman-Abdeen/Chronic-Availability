using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Chronic_Availability
{
    public class SiteStatus
    {
        public string NodeID { get; set; }
        public string controller { get; set; }
        public string Status { get; set; }
        public string MarkedROT { get; set; }
        public string subcategory { get; set; }
        public string area { get; set; }
        public string Tier { get; set; }
        public string vendor { get; set; }
        public string ServiceType { get; set; }
        public Boolean Is2G { get; set; }
        public string SiteID { get; set; }
        public string NumberOfActiveCells { get; set; }

    }
}
