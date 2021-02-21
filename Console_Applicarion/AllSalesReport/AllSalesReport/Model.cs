using System;
using System.Collections.Generic;

namespace AllSalesReport
{
    enum Categories
    {
        Open_Pool = 1,
        Dedicated_Loops = 2,
        Inside_Sales = 3,
        Closed_Loops = 4,
        Meat = 5,
        Others = 6
    }

    enum DataType
    {
        Forecast = -4,
        Per_Forecast = -3,
        Budget = -2,
        Per_Budget = -1,
        LY = 0,
        TY = 1,
        Per_Chg = 2
    }

    enum SalesRepCodes
    {
        // East
        Tony_Mosco = 14,
        Dana_Embry = 16,
        Susan_Rivera = 17,
        Brandon_Osteen = 21,
        Travis_White = 22,
        Frank_Niolet = 24,
        // West
        Sarah_Gonzalez = 3,
        Gabriel_Grijalva = 15,
        Paul_Pederson = 20,
        Connie_Ceniceros = 23,
        Lisa_Dennis = 25,
        // All
        Eric_Biddiscombe = 26
    }

    interface ILine
    {
        int Data_Type { get; set; }
        double M1 { get; set; }
        double M2 { get; set; }
        double M3 { get; set; }
        double M4 { get; set; }
        double M5 { get; set; }
        double M6 { get; set; }
        double M7 { get; set; }
        double M8 { get; set; }
        double M9 { get; set; }
        double M10 { get; set; }
        double M11 { get; set; }
        double M12 { get; set; }
        double Total { get; set; }
    }

    interface IOfficeSaleDesc
    {
        string Office_Sale_Desc { get; set; }
    }

    interface ILineOfficeSaleDesc : ILine, IOfficeSaleDesc
    {
    }

    interface ILineLight
    {
        string Loop_Desc { get; set; }
        string Commodity_Desc { get; set; }
    }

    class TotalLine : ILine
    {
        public string Name { get; set; }
        public int Data_Type { get; set; }
        public double M1 { get; set; }
        public double M2 { get; set; }
        public double M3 { get; set; }
        public double M4 { get; set; }
        public double M5 { get; set; }
        public double M6 { get; set; }
        public double M7 { get; set; }
        public double M8 { get; set; }
        public double M9 { get; set; }
        public double M10 { get; set; }
        public double M11 { get; set; }
        public double M12 { get; set; }
        public double Total { get; set; }
    }

    class DetailsLine : ILine, IOfficeSaleDesc, ILineOfficeSaleDesc, ILineLight
    {
        public int Is_Virtual { get; set; }
        public int Customer_Code { get; set; }
        public string Customer_Name { get; set; }
        public int Office_Sale_Code { get; set; }
        public string Office_Sale_Desc { get; set; }
        public int Product_Code { get; set; }
        public string Product_Name { get; set; }
        public int Data_Type { get; set; }
        public string Data_Desc { get; set; }
        public string Loop_Desc { get; set; }
        public string Commodity_Desc { get; set; }
        public double M1 { get; set; }
        public double M2 { get; set; }
        public double M3 { get; set; }
        public double M4 { get; set; }
        public double M5 { get; set; }
        public double M6 { get; set; }
        public double M7 { get; set; }
        public double M8 { get; set; }
        public double M9 { get; set; }
        public double M10 { get; set; }
        public double M11 { get; set; }
        public double M12 { get; set; }
        public double Total { get; set; }
        public long Row_No_Customer_Product_Data_Type { get; set; }
        public int Is_East { get; set; }
        public int Is_West { get; set; }
        public int Current_Year { get; set; }
        public int Current_Month { get; set; }
    }

    class CustomerSummaryLine : ILine, IOfficeSaleDesc, ILineOfficeSaleDesc
    {
        public int Is_Virtual { get; set; }
        public int Customer_Code { get; set; }
        public string Customer_Name { get; set; }
        public int Office_Sale_Code { get; set; }
        public string Office_Sale_Desc { get; set; }
        public int Data_Type { get; set; }
        public string Data_Desc { get; set; }
        public double M1 { get; set; }
        public double M2 { get; set; }
        public double M3 { get; set; }
        public double M4 { get; set; }
        public double M5 { get; set; }
        public double M6 { get; set; }
        public double M7 { get; set; }
        public double M8 { get; set; }
        public double M9 { get; set; }
        public double M10 { get; set; }
        public double M11 { get; set; }
        public double M12 { get; set; }
        public double Total { get; set; }
        public long Row_No_Customer_Data_Type { get; set; }
        public int Is_East { get; set; }
        public int Is_West { get; set; }
        public int Current_Year { get; set; }
        public int Current_Month { get; set; }
    }

    class CustomerSummaryLineLight : ILine, IOfficeSaleDesc, ILineOfficeSaleDesc, ILineLight
    {
        public int Is_Virtual { get; set; }
        public int Customer_Code { get; set; }
        public string Customer_Name { get; set; }
        public int Office_Sale_Code { get; set; }
        public string Office_Sale_Desc { get; set; }
        public int Data_Type { get; set; }
        public string Data_Desc { get; set; }
        public string Loop_Desc { get; set; }
        public string Commodity_Desc { get; set; }
        public double M1 { get; set; }
        public double M2 { get; set; }
        public double M3 { get; set; }
        public double M4 { get; set; }
        public double M5 { get; set; }
        public double M6 { get; set; }
        public double M7 { get; set; }
        public double M8 { get; set; }
        public double M9 { get; set; }
        public double M10 { get; set; }
        public double M11 { get; set; }
        public double M12 { get; set; }
        public double Total { get; set; }
        public long Row_No_Customer_Data_Type { get; set; }
        public int Is_East { get; set; }
        public int Is_West { get; set; }
        public int Current_Year { get; set; }
        public int Current_Month { get; set; }
    }

    class LYVsTYBySizeLine : ILine
    {
        public int Is_Virtual { get; set; }
        public int Product_Code { get; set; }
        public string Product_Name { get; set; }
        public int Data_Type { get; set; }
        public string Data_Desc { get; set; }
        public double M1 { get; set; }
        public double M2 { get; set; }
        public double M3 { get; set; }
        public double M4 { get; set; }
        public double M5 { get; set; }
        public double M6 { get; set; }
        public double M7 { get; set; }
        public double M8 { get; set; }
        public double M9 { get; set; }
        public double M10 { get; set; }
        public double M11 { get; set; }
        public double M12 { get; set; }
        public double Total { get; set; }
        public long Row_No_Product_Data_Type { get; set; }
        public int Current_Year { get; set; }
        public int Current_Month { get; set; }
    }

    class RelCustomerCategory
    {
        public int Customer_Code { get; set; }
        public string Customer_Name { get; set; }
        public int Category_Code { get; set; }
        public string Category_Desc { get; set; }
    }

    class SalesRep
    {
        public int Office_Sale_Code { get; set; }
        public string Office_Sale_Desc { get; set; }
        public int Is_East { get; set; }
        public int Is_West { get; set; }
    }

    class Model
    {
        public IEnumerable<DetailsLine> Details { get; set; }
        public IEnumerable<DetailsLine> EastDetails { get; set; }
        public IEnumerable<DetailsLine> WestDetails { get; set; }
        public IEnumerable<CustomerSummaryLine> CustomerSummary { get; set; }
        public IEnumerable<CustomerSummaryLine> EastCustomerSummary { get; set; }
        public IEnumerable<CustomerSummaryLine> WestCustomerSummary { get; set; }
        public IEnumerable<CustomerSummaryLineLight> CustomerSummaryLight { get; set; }
        public IEnumerable<LYVsTYBySizeLine> LyVsTYBySize { get; set; }
        public IEnumerable<LYVsTYBySizeLine> EastLyVsTYBySize { get; set; }
        public IEnumerable<LYVsTYBySizeLine> WestLyVsTYBySize { get; set; }
        public IEnumerable<RelCustomerCategory> RelCustomerCategory { get; set; }
        public IEnumerable<SalesRep> SalesReps { get; set; }
        public IEnumerable<RelCustomerCategory> CustomersOpenPool { get; set; }
        public IEnumerable<RelCustomerCategory> CustomersDedicatedLoops { get; set; }
        public IEnumerable<RelCustomerCategory> CustomersInsideSales { get; set; }
        public IEnumerable<RelCustomerCategory> CustomersClosedLoops { get; set; }
        public IEnumerable<RelCustomerCategory> CustomersMeat { get; set; }
        public IEnumerable<RelCustomerCategory> CustomersOthers { get; set; }
        public IEnumerable<SalesRep> EastSalesReps { get; set; }
        public IEnumerable<SalesRep> WestSalesReps { get; set; }
    }
}
