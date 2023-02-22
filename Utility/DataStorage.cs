using ExportToExcelSample.Models;
using System;
using System.Collections.Generic;

namespace ExportToExcelSample.Utility
{
    public static class DataStorage
    {
        public static List<Market> GetAllMarkets() =>
            new()
            {
                new Market { MarketId=1, Name="Dallas - Ft. Worth", TabColor="Grey" }, // NOTE: '/' not allowed in sheet name
                new Market { MarketId=2, Name="Austin", TabColor="Green" },
                new Market { MarketId=3, Name="DFW 2", TabColor="Blue" },
            };

        public static List<Available> GetAllAvailable() =>
            new()
            {
                new Available { MarketId = 1, Address="9749 Clifford Dr", City="Dallas", State="TX", TotalSf=109200, ClearHt=30, Doors=28, AvailableDate=DateTime.Now, ListingRep="Tracy Bertman", RepEmail="TBertman@GMail.com"},
                new Available { MarketId = 1, Address="3150 Marquis Dr", City="Garland", State="TX", TotalSf=355071, ClearHt=32, Doors=128, AvailableDate=new DateTime(2023, 12, 23), ListingRep="Tracy Bertman", RepEmail="TBertman@GMail.com"},
                new Available { MarketId = 1, Address="2601 West Bethel Rd", City="Dallas", State="TX", TotalSf=1008176, ClearHt=42, Doors=70, AvailableDate=DateTime.Now, ListingRep="Tracy Bertman", RepEmail="TBertman@GMail.com"},
                new Available { MarketId = 1, Address="Strandridge Dr & Memorial Dr", City="The Colony", State="TX", TotalSf=1107848, ClearHt=32, Doors=0, AvailableDate=new DateTime(2023, 03, 26), ListingRep="Tracy Bertman", RepEmail = "TBertman@GMail.com"},
                new Available { MarketId = 1, Address="9233 Waterford Centre Blvd", City="Dallas", State="TX", TotalSf=508364, ClearHt=28, Doors=4, AvailableDate=DateTime.Now, ListingRep="Tracy Bertman", RepEmail = "TBertman@GMail.com"},
                new Available { MarketId = 1, Address="363 Imagine Ave", City="Dallas", State="TX", TotalSf=230913, ClearHt=24, Doors=22, AvailableDate=new DateTime(2023, 04, 06), ListingRep="Tracy Bertman", RepEmail = "TBertman@GMail.com"},

                new Available { MarketId = 2, Address="912521 Harris Branch Parkway", City="Austin", State="TX", TotalSf=654080, ClearHt=32, Doors=141, AvailableDate=new DateTime(2023, 05, 17), ListingRep="Richard Myers", RepEmail="RMyers@Gmail.com"},
                new Available { MarketId = 2, Address="2111 A.W. Grimes Blvd", City="Austin", State="TX", TotalSf=367887, ClearHt=32, Doors=95, AvailableDate=DateTime.Now, ListingRep="Richard Myers", RepEmail="RMyers@Gmail.com"},
                new Available { MarketId = 2, Address="2100 Chisolm Trl", City="Round Rock", State="TX", TotalSf=663267, ClearHt=34, Doors=148, AvailableDate=new DateTime(2023, 06, 25), ListingRep="Richard Myers", RepEmail="RMyers@Gmail.com"},
                new Available { MarketId = 2, Address="912521 Harris Branch Parkway", City="Austin", State="TX", TotalSf=654080, ClearHt=32, Doors=141, AvailableDate=new DateTime(2023, 05, 17), ListingRep="Richard Myers", RepEmail="RMyers@Gmail.com"},
                new Available { MarketId = 2, Address="2111 A.W. Grimes Blvd", City="Austin", State="TX", TotalSf=367887, ClearHt=32, Doors=95, AvailableDate=DateTime.Now, ListingRep="Richard Myers", RepEmail="RMyers@Gmail.com"},
                new Available { MarketId = 2, Address="2100 Chisolm Trl", City="Round Rock", State="TX", TotalSf=663267, ClearHt=34, Doors=148, AvailableDate=new DateTime(2023, 06, 25), ListingRep="Richard Myers", RepEmail="RMyers@Gmail.com"},
                new Available { MarketId = 2, Address="912521 Harris Branch Parkway", City="Austin", State="TX", TotalSf=654080, ClearHt=32, Doors=141, AvailableDate=new DateTime(2023, 05, 17), ListingRep="Richard Myers", RepEmail="RMyers@Gmail.com"},
                new Available { MarketId = 2, Address="2111 A.W. Grimes Blvd", City="Austin", State="TX", TotalSf=367887, ClearHt=32, Doors=95, AvailableDate=DateTime.Now, ListingRep="Richard Myers", RepEmail="RMyers@Gmail.com"},
                new Available { MarketId = 2, Address="2100 Chisolm Trl", City="Round Rock", State="TX", TotalSf=663267, ClearHt=34, Doors=148, AvailableDate=new DateTime(2023, 06, 25), ListingRep="Richard Myers", RepEmail="RMyers@Gmail.com"},

                new Available { MarketId = 3, Address="9749 Clifford Dr", City="Dallas", State="TX", TotalSf=109200, ClearHt=30, Doors=28, AvailableDate=DateTime.Now, ListingRep="Tracy Bertman", RepEmail="TBertman@GMail.com"},
                new Available { MarketId = 3, Address="3150 Marquis Dr", City="Garland", State="TX", TotalSf=355071, ClearHt=32, Doors=128, AvailableDate=new DateTime(2023, 03, 23), ListingRep="Tracy Bertman", RepEmail="TBertman@GMail.com"},
                new Available { MarketId = 3, Address="2601 West Bethel Rd", City="Dallas", State="TX", TotalSf=1008176, ClearHt=42, Doors=70, AvailableDate=DateTime.Now, ListingRep="Tracy Bertman", RepEmail="TBertman@GMail.com"},
                new Available { MarketId = 3, Address="Strandridge Dr & Memorial Dr", City="The Colony", State="TX", TotalSf=1107848, ClearHt=32, Doors=0, AvailableDate=new DateTime(2023, 03, 26), ListingRep="Tracy Bertman", RepEmail="TBertman@GMail.com"},
                new Available { MarketId = 3, Address="9233 Waterford Centre Blvd", City="Dallas", State="TX", TotalSf=508364, ClearHt=28, Doors=4, AvailableDate=DateTime.Now, ListingRep="Tracy Bertman", RepEmail="TBertman@GMail.com"},
                new Available { MarketId = 3, Address="363 Imagine Ave", City="Dallas", State="TX", TotalSf=230913, ClearHt=24, Doors=22, AvailableDate=new DateTime(2023, 04, 23), ListingRep="Tracy Bertman", RepEmail="TBertman@GMail.com"},
            };
    }
}
