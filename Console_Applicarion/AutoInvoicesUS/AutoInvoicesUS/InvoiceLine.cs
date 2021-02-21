using System;

namespace AutoInvoicesUS
{
    class InvoiceLine
    {
        public int? Customer_Code { get; set; }
        public string Customer_Name { get; set; }
        public int? Site_Code { get; set; }
        public string Site_Name { get; set; }
        public DateTime? Value_Date { get; set; }
        public DateTime? Delivery_Date { get; set; }
        public int Numerator { get; set; }
        public int? Has_Scanned_Notes { get; set; }
        public int? Product_Code { get; set; }
        public string Product_Name { get; set; }
        public int? Quantity { get; set; }
        public decimal? Price { get; set; }
        public decimal? Total { get; set; }
        public int? From_Site_Code { get; set; }
        public string From_Site_Name { get; set; }
        public string Reference { get; set; }
        public string Note { get; set; }
        public string Sap_Customer { get; set; }
        public string Sap_Site { get; set; }
        public string Sap_Product { get; set; }
        public int? Warehouse_Code { get; set; }
        public char? Price_Type_Code_Char { get; set; }
        public short? Invoice_Per_Site { get; set; }
        public int? Operation_Code { get; set; }
        public string Operation_Desc { get; set; }
        public int? Order_Numerator { get; set; }
        public int? Haulier_Code { get; set; }
        public string Haulier_Name { get; set; }
        public int? FSC_Mileage { get; set; }
        public short? Is_FSC { get; set; }
        public short? Internal_Use { get; set; }
        public string Internal_Use_Desc { get; set; }
        public decimal? Deduction { get; set; }
        public string Tax_Exemption_PTE { get; set; }
        public int? Is_Manual_Calculation { get; set; }
        public decimal? Sales_Tax_State { get; set; }
        public decimal? Sales_Tax_County { get; set; }
        public decimal? Sales_Tax_City { get; set; }
        public decimal? Tax_Rates_Percent { get; set; }
        public decimal? Final_Sales_Tax { get; set; }
        public decimal? Sales_Tax_Amount { get; set; }

        public string Comments { get; set; }
    }
}
