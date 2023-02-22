using System;

namespace ExportToExcelSample.Models 
{
    public class Available 
    { 
        public int MarketId { get; set; }
        public string Address { get; set; }
        public string City { get; set; } 
        public string State { get; set; } 
        public int TotalSf { get; set; }
        public int ClearHt { get; set; }
        public int Doors { get; set; }
        public DateTime AvailableDate { get; set; }
        public string ListingRep { get; set; }
        public string RepEmail { get; set; }
    }
}