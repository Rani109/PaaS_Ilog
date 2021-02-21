using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using iTextSharp.text;
using iTextSharp.text.pdf;
using MathNet.Numerics.LinearAlgebra;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace AutoInvoicesUS
{
    static class ExportHelper
    {
        private static readonly string[] ExportHeaders = new string[] { 
            "Customer", 
            "Site", 
            "Value Date", 
            "Delivery Date",
            "Numerator", 
            "Order",
            "Carrier",
            "Product", 
            "Quantity", 
            "Price",
            "Total",
            "Customer Internal Use", 
            "Sales Tax %", 
            "Deduction %",
            "Final Sales Tax %", 
            "Sales Tax Amount",
            "Comments",
            "Tax Exemption (P/T/E)", 
            "From Site", 
            "State Sales Tax",
            "County Sales Tax",
            "City Sales Tax"
        };

        private static readonly bool[] ExportHorizontalAlignmentRight = new bool[] { 
            false,
            false,
            true,
            true,
            false,
            false,
            false,
            false,
            true,
            true,
            true,
            false,
            true,
            true,
            true,
            true,
            false,
            false,
            false,
            true,
            true,
            true
        };

        private static readonly int ExportColumnsCount = ExportHeaders.Length;

        private const string REPORT_NAME = "Upload Sales invoice";
        private const string SEARCH_CRITERIA_1 = "Numerator: {0}, Packer To Packer: No, Grouping: BOL, ";
        private const string SEARCH_CRITERIA_2 = "Total report lines: {0}";

        #region PDF

        public static class PDFExportHelper
        {
            private static readonly BaseColor HeaderColor = new BaseColor(230, 230, 230);
            private static readonly BaseColor SearchCriteriaColor = new BaseColor(230, 230, 230);
            private static readonly BaseColor SearchCriteria2FontColor = new BaseColor(51, 153, 102);
            private static readonly BaseColor OutlineHeaderColor = new BaseColor(198, 224, 180);
            private static readonly BaseColor OutlineHeaderFontColor = new BaseColor(63, 63, 63);
            private static readonly BaseColor OutlineColor = new BaseColor(245, 245, 245);
            private static readonly BaseColor SumRowColor = new BaseColor(143, 209, 49);
            private static readonly BaseColor Green = new BaseColor(51, 153, 102);

            private const float STATIC_FONT_SIZE = 10F;

            private static readonly iTextSharp.text.Font FontDefault = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, STATIC_FONT_SIZE);
            private static readonly iTextSharp.text.Font FontDefaultBold = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, STATIC_FONT_SIZE, iTextSharp.text.Font.BOLD);
            private static readonly iTextSharp.text.Font FontOutline = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, STATIC_FONT_SIZE, iTextSharp.text.Font.ITALIC, OutlineHeaderFontColor);
            private static readonly iTextSharp.text.Font FontReportName = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, STATIC_FONT_SIZE + 3.5F, iTextSharp.text.Font.BOLD);
            private static readonly iTextSharp.text.Font FontSearchCriteria2 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, STATIC_FONT_SIZE, iTextSharp.text.Font.NORMAL, SearchCriteria2FontColor);
            private static readonly iTextSharp.text.Font FontDefaultBoldGreen = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, STATIC_FONT_SIZE, iTextSharp.text.Font.BOLD, Green);

            private static readonly iTextSharp.text.Font FontDefaultHebrew;
            private static readonly iTextSharp.text.Font FontDefaultBoldHebrew;
            private static readonly iTextSharp.text.Font FontOutlineHebrew;
            private static readonly iTextSharp.text.Font FontReportNameHebrew;
            private static readonly iTextSharp.text.Font FontSearchCriteria2Hebrew;
            private static readonly iTextSharp.text.Font FontDefaultBoldGreenHebrew = null;

            private const char CurrencySymbol = '£';
            private const string ShortDateFormat = "{0:d}";
            private const string LongDateFormat = "{0:G}";
            private const string PrintingFormat = "{0:g}";
            private const string IntFormat = "{0:#,##0}";
            private const string DoubleFormat = "{0:#,##0.000000}";
            private const string CurrencyIntFormat = "£{0:#,##0}";
            private const string CurrencyDoubleFormat = "£{0:#,##0.000000}";
            private const string PercentageIntFormat = "{0:0%}";
            private const string PercentageDoubleFormat = "{0:0.000000%}";
            private static readonly CultureInfo cultureInfo = CultureInfo.CreateSpecificCulture(string.Empty);

            private const string PDF_METADATA_AUTHOR = "Polymer Logistics";
            private const string PDF_METADATA_CREATOR = "Polymer Logistics";
            private const string PDF_METADATA_KEYWORDS = "Polymer Logistics";

            static PDFExportHelper()
            {
                cultureInfo.DateTimeFormat.TimeSeparator = ":";
                cultureInfo.DateTimeFormat.DateSeparator = "/";
                cultureInfo.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy"; // d
                cultureInfo.DateTimeFormat.FullDateTimePattern = "dd/MM/yyyy HH\\:mm\\:ss"; // G
                cultureInfo.NumberFormat.CurrencySymbol = "£";
                cultureInfo.NumberFormat.CurrencyGroupSeparator = ",";
                cultureInfo.NumberFormat.CurrencyDecimalSeparator = ".";
                cultureInfo.NumberFormat.NumberGroupSeparator = ",";
                cultureInfo.NumberFormat.NumberDecimalSeparator = ".";

                BaseFont hebrewBaseFont = null;
                string fontPath = "c:/windows/fonts/";
                if (System.IO.File.Exists(fontPath + "David.ttf"))
                    hebrewBaseFont = BaseFont.CreateFont(fontPath + "David.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                else if (System.IO.File.Exists(fontPath + "Gisha.ttf"))
                    hebrewBaseFont = BaseFont.CreateFont(fontPath + "Gisha.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

                if (hebrewBaseFont != null)
                {
                    FontDefaultHebrew = new iTextSharp.text.Font(hebrewBaseFont, STATIC_FONT_SIZE);
                    FontDefaultBoldHebrew = new iTextSharp.text.Font(hebrewBaseFont, STATIC_FONT_SIZE, iTextSharp.text.Font.BOLD);
                    FontOutlineHebrew = new iTextSharp.text.Font(hebrewBaseFont, STATIC_FONT_SIZE, iTextSharp.text.Font.ITALIC, OutlineHeaderFontColor);
                    FontReportNameHebrew = new iTextSharp.text.Font(hebrewBaseFont, STATIC_FONT_SIZE + 3.5F, iTextSharp.text.Font.BOLD);
                    FontSearchCriteria2Hebrew = new iTextSharp.text.Font(hebrewBaseFont, STATIC_FONT_SIZE, iTextSharp.text.Font.NORMAL, SearchCriteria2FontColor);
                    FontDefaultBoldGreenHebrew = new iTextSharp.text.Font(hebrewBaseFont, STATIC_FONT_SIZE, iTextSharp.text.Font.BOLD, Green);
                }
                else
                {
                    FontDefaultHebrew = FontDefault;
                    FontDefaultBoldHebrew = FontDefaultBold;
                    FontOutlineHebrew = FontOutline;
                    FontReportNameHebrew = FontReportName;
                    FontSearchCriteria2Hebrew = FontSearchCriteria2;
                    FontDefaultBoldGreenHebrew = FontDefaultBoldGreen;
                }
            }

            private static PdfPTable GetPDFTable(int numColumns)
            {
                PdfPTable table = new PdfPTable(numColumns);
                table.HorizontalAlignment = Rectangle.ALIGN_CENTER;
                table.WidthPercentage = 100F;
                return table;
            }

            private class PdfPageEvent : PdfPageEventHelper
            {
                public Document document { get; set; }
                public DateTime now { get; set; }
                public List<string> columnNames { get; set; }
                public string PrintingFormat { get; set; }
                public float[] relativeWidths { get; set; }
                public float cellPadding { get; set; }
                public Func<string, bool> IsLTR { get; set; }

                public int linesCount { get; set; }
                public int numerator { get; set; }

                public iTextSharp.text.Font FontDefault { get; set; }
                public iTextSharp.text.Font FontDefaultBold { get; set; }
                public iTextSharp.text.Font FontOutline { get; set; }
                public iTextSharp.text.Font FontReportName { get; set; }
                public iTextSharp.text.Font FontSearchCriteria2 { get; set; }

                public iTextSharp.text.Font FontDefaultHebrew { get; set; }
                public iTextSharp.text.Font FontDefaultBoldHebrew { get; set; }
                public iTextSharp.text.Font FontOutlineHebrew { get; set; }
                public iTextSharp.text.Font FontReportNameHebrew { get; set; }
                public iTextSharp.text.Font FontSearchCriteria2Hebrew { get; set; }

                private PdfPTable table;

                public override void OnStartPage(PdfWriter writer, Document document)
                {
                    if (table == null)
                        BuildHeader();
                    document.Add(table);
                }

                private void BuildHeader()
                {
                    table = GetPDFTable(ExportColumnsCount);

                    // logo
                    var image = iTextSharp.text.Image.GetInstance(Assembly.GetExecutingAssembly().GetManifestResourceStream("AutoInvoicesUS.company_logo_big.gif"));

                    PdfPCell cell = new PdfPCell(image);
                    cell.Colspan = ExportColumnsCount;
                    cell.HorizontalAlignment = Rectangle.ALIGN_LEFT;
                    cell.VerticalAlignment = Rectangle.ALIGN_MIDDLE;
                    cell.Border = Rectangle.NO_BORDER;
                    cell.NoWrap = false;
                    cell.Padding = 0F;
                    table.AddCell(cell);

                    // printing date
                    cell = new PdfPCell(new Phrase(string.Format(PrintingFormat, now, cultureInfo), FontDefault));
                    cell.Colspan = ExportColumnsCount;
                    cell.HorizontalAlignment = Rectangle.ALIGN_RIGHT;
                    cell.VerticalAlignment = Rectangle.ALIGN_MIDDLE;
                    cell.Border = Rectangle.NO_BORDER;
                    cell.NoWrap = false;
                    cell.Padding = cellPadding;
                    table.AddCell(cell);

                    // report name
                    bool isLTR = IsLTR(REPORT_NAME);
                    cell = new PdfPCell(new Phrase(REPORT_NAME, (isLTR ? FontReportName : FontReportNameHebrew)));
                    cell.RunDirection = (isLTR ? PdfWriter.RUN_DIRECTION_LTR : PdfWriter.RUN_DIRECTION_RTL);
                    cell.Colspan = ExportColumnsCount;
                    cell.HorizontalAlignment = Rectangle.ALIGN_CENTER;
                    cell.VerticalAlignment = Rectangle.ALIGN_MIDDLE;
                    cell.Border = Rectangle.TOP_BORDER | Rectangle.LEFT_BORDER | Rectangle.RIGHT_BORDER;
                    cell.BackgroundColor = HeaderColor;
                    cell.NoWrap = false;
                    cell.Padding = cellPadding;
                    cell.PaddingBottom = 0F;
                    table.AddCell(cell);

                    // search criteria
                    string searchCriteria1 = string.Format(SEARCH_CRITERIA_1, numerator);
                    string searchCriteriaTotal = string.Format(SEARCH_CRITERIA_2, linesCount);

                    isLTR = true;
                    if (string.IsNullOrEmpty(searchCriteria1) == false)
                        isLTR = IsLTR(searchCriteria1);
                    Phrase phrase = new Phrase();
                    if (string.IsNullOrEmpty(searchCriteria1) == false)
                        phrase.Add(new Phrase(searchCriteria1, (isLTR ? FontDefault : FontDefaultHebrew)));
                    if (string.IsNullOrEmpty(searchCriteriaTotal) == false)
                        phrase.Add(new Phrase(searchCriteriaTotal, (isLTR ? FontSearchCriteria2 : FontSearchCriteria2Hebrew)));
                    cell = new PdfPCell(phrase);
                    cell.RunDirection = (isLTR ? PdfWriter.RUN_DIRECTION_LTR : PdfWriter.RUN_DIRECTION_RTL);
                    cell.Colspan = ExportColumnsCount;
                    cell.HorizontalAlignment = Rectangle.ALIGN_CENTER;
                    cell.VerticalAlignment = Rectangle.ALIGN_MIDDLE;
                    cell.Border = Rectangle.BOTTOM_BORDER | Rectangle.LEFT_BORDER | Rectangle.RIGHT_BORDER;
                    cell.BackgroundColor = SearchCriteriaColor;
                    cell.NoWrap = false;
                    cell.Padding = cellPadding;
                    table.AddCell(cell);

                    // headers
                    for (int i = 0; i < ExportColumnsCount; i++)
                    {
                        string header = columnNames[i];
                        isLTR = IsLTR(header);
                        cell = new PdfPCell(new Phrase(header, (isLTR ? FontDefault : FontDefaultHebrew)));
                        cell.RunDirection = (isLTR ? PdfWriter.RUN_DIRECTION_LTR : PdfWriter.RUN_DIRECTION_RTL);
                        cell.HorizontalAlignment = Rectangle.ALIGN_CENTER;
                        cell.VerticalAlignment = Rectangle.ALIGN_MIDDLE;
                        cell.Border = Rectangle.BOX;
                        cell.BackgroundColor = HeaderColor;
                        cell.NoWrap = false;
                        cell.Padding = cellPadding;
                        table.AddCell(cell);
                    }

                    table.SetWidths(relativeWidths);
                }
            }

            public static byte[] GetPDF(int numerator, string customerName, IEnumerable<InvoiceLine> invoiceLines, IEnumerable<FSCMileageLine> mileageLines, DateTime now)
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    using (Document document = new Document())
                    {
                        using (PdfWriter writer = PdfWriter.GetInstance(document, ms))
                        {
                            int invoiceLinesCount = invoiceLines.Count();
                            int mileageLinesCount = mileageLines.Count();
                            int linesCount = invoiceLinesCount + mileageLinesCount;

                            string[][] tableValues = new string[linesCount + 1][];
                            int tableRowIndex = 0;

                            // invoice rows
                            foreach (var line in invoiceLines)
                            {
                                string Customer_Name = line.Customer_Name ?? string.Empty;
                                string Site_Name = line.Site_Name ?? string.Empty;

                                string Value_Date = string.Empty;
                                if (line.Value_Date != null)
                                    Value_Date = string.Format(ShortDateFormat, line.Value_Date.Value, cultureInfo);

                                string Delivery_Date = string.Empty;
                                if (line.Delivery_Date != null)
                                    Delivery_Date = string.Format(ShortDateFormat, line.Delivery_Date.Value, cultureInfo);

                                string Numerator = line.Numerator.ToString();

                                string Order_Numerator = string.Empty;
                                if (line.Order_Numerator != null)
                                    Order_Numerator = line.Order_Numerator.Value.ToString();

                                string Haulier_Name = line.Haulier_Name ?? string.Empty;
                                string Product_Name = line.Product_Name ?? string.Empty;

                                string Quantity = string.Empty;
                                if (line.Quantity != null)
                                    Quantity = string.Format(IntFormat, line.Quantity.Value, cultureInfo);

                                string Price = string.Empty;
                                if (line.Price != null)
                                    Price = string.Format(CurrencyDoubleFormat, line.Price.Value, cultureInfo);

                                string Total = string.Empty;
                                if (line.Total != null)
                                    Total = string.Format(CurrencyDoubleFormat, line.Total.Value, cultureInfo);

                                string Internal_Use_Desc = line.Internal_Use_Desc ?? string.Empty;

                                string Tax_Rates_Percent = string.Empty;
                                if (line.Tax_Rates_Percent != null)
                                    Tax_Rates_Percent = string.Format(PercentageDoubleFormat, line.Tax_Rates_Percent.Value, cultureInfo);

                                string Deduction = string.Empty;
                                if (line.Deduction != null)
                                    Deduction = string.Format(PercentageDoubleFormat, line.Deduction.Value, cultureInfo);

                                string Final_Sales_Tax = string.Empty;
                                if (line.Final_Sales_Tax != null)
                                    Final_Sales_Tax = string.Format(PercentageDoubleFormat, line.Final_Sales_Tax.Value, cultureInfo);

                                string Sales_Tax_Amount = string.Empty;
                                if (line.Sales_Tax_Amount != null)
                                    Sales_Tax_Amount = string.Format(CurrencyDoubleFormat, line.Sales_Tax_Amount.Value, cultureInfo);

                                string Comments = line.Comments ?? string.Empty;
                                string Tax_Exemption_PTE = line.Tax_Exemption_PTE ?? string.Empty;
                                string From_Site_Name = line.From_Site_Name ?? string.Empty;

                                string Sales_Tax_State = string.Empty;
                                if (line.Sales_Tax_State != null)
                                    Sales_Tax_State = string.Format(DoubleFormat, line.Sales_Tax_State.Value, cultureInfo);

                                string Sales_Tax_County = string.Empty;
                                if (line.Sales_Tax_County != null)
                                    Sales_Tax_County = string.Format(DoubleFormat, line.Sales_Tax_County.Value, cultureInfo);

                                string Sales_Tax_City = string.Empty;
                                if (line.Sales_Tax_City != null)
                                    Sales_Tax_City = string.Format(DoubleFormat, line.Sales_Tax_City.Value, cultureInfo);

                                tableValues[tableRowIndex] = new string[]
                                {
                                    Customer_Name,
                                    Site_Name,
                                    Value_Date,
                                    Delivery_Date,
                                    Numerator,
                                    Order_Numerator,
                                    Haulier_Name,
                                    Product_Name,
                                    Quantity,
                                    Price,
                                    Total,
                                    Internal_Use_Desc,
                                    Tax_Rates_Percent,
                                    Deduction,
                                    Final_Sales_Tax,
                                    Sales_Tax_Amount,
                                    Comments,
                                    Tax_Exemption_PTE,
                                    From_Site_Name,
                                    Sales_Tax_State,
                                    Sales_Tax_County,
                                    Sales_Tax_City
                                };

                                tableRowIndex++;
                            }

                            // mileage rows
                            foreach (var line in mileageLines)
                            {
                                string Customer_Name = customerName ?? string.Empty;
                                string Site_Name = string.Empty;
                                string Value_Date = string.Empty;
                                string Delivery_Date = string.Empty;
                                string Numerator = line.Numerator.ToString();
                                string Order_Numerator = string.Empty;
                                string Haulier_Name = string.Empty;
                                string Product_Name = "Fuel Surcharge";
                                string Quantity = string.Format(IntFormat, line.Priority_FSC_Mileage, cultureInfo);
                                string Price = string.Format(CurrencyDoubleFormat, line.Priority_FSC_Price, cultureInfo);
                                string Total = string.Format(CurrencyDoubleFormat, line.Priority_FSC_Price, cultureInfo);
                                string Internal_Use_Desc = string.Empty;
                                string Tax_Rates_Percent = string.Empty;
                                string Deduction = string.Empty;
                                string Final_Sales_Tax = string.Empty;
                                string Sales_Tax_Amount = string.Empty;
                                string Comments = string.Empty;
                                string Tax_Exemption_PTE = string.Empty;
                                string From_Site_Name = string.Empty;
                                string Sales_Tax_State = string.Empty;
                                string Sales_Tax_County = string.Empty;
                                string Sales_Tax_City = string.Empty;

                                tableValues[tableRowIndex] = new string[]
                                {
                                    Customer_Name,
                                    Site_Name,
                                    Value_Date,
                                    Delivery_Date,
                                    Numerator,
                                    Order_Numerator,
                                    Haulier_Name,
                                    Product_Name,
                                    Quantity,
                                    Price,
                                    Total,
                                    Internal_Use_Desc,
                                    Tax_Rates_Percent,
                                    Deduction,
                                    Final_Sales_Tax,
                                    Sales_Tax_Amount,
                                    Comments,
                                    Tax_Exemption_PTE,
                                    From_Site_Name,
                                    Sales_Tax_State,
                                    Sales_Tax_County,
                                    Sales_Tax_City
                                };

                                tableRowIndex++;
                            }

                            // sum row
                            int invoiceLinesSumQuantity = invoiceLines.Sum(x => x.Quantity) ?? 0;
                            int mileageLinesSumQuantity = mileageLines.Sum(x => x.Priority_FSC_Mileage);
                            int sumQuantity = invoiceLinesSumQuantity + mileageLinesSumQuantity;

                            decimal invoiceLinesSumTotal = invoiceLines.Sum(x => x.Total) ?? 0;
                            decimal mileageLinesSumTotal = mileageLines.Sum(x => x.Priority_FSC_Price);
                            decimal sumTotal = invoiceLinesSumTotal + mileageLinesSumTotal;

                            tableValues[tableRowIndex] = new string[]
                            {
                                string.Empty,
                                string.Empty,
                                string.Empty,
                                string.Empty,
                                string.Empty,
                                string.Empty,
                                string.Empty,
                                string.Empty,
                                string.Format(IntFormat, sumQuantity, cultureInfo),
                                string.Empty,
                                string.Format(CurrencyDoubleFormat, sumTotal, cultureInfo),
                                string.Empty,
                                string.Empty,
                                string.Empty,
                                string.Empty,
                                string.Empty,
                                string.Empty,
                                string.Empty,
                                string.Empty,
                                string.Empty,
                                string.Empty,
                                string.Empty
                            };

                            // relative widths
                            BaseFont baseFont = FontDefault.GetCalculatedBaseFont(true);
                            BaseFont baseFontHebrew = FontDefaultHebrew.GetCalculatedBaseFont(true);

                            float minColumnWidth = 32F;
                            float[] relativeWidths = Enumerable.Repeat<float>(minColumnWidth, ExportColumnsCount).ToArray();
                            foreach (var row in tableValues)
                            {
                                for (int columnIndex = 0; columnIndex < row.Length; columnIndex++)
                                {
                                    string valueStr = row[columnIndex];

                                    if (string.IsNullOrEmpty(valueStr) == false)
                                    {
                                        bool isLTR = IsLTR(valueStr);
                                        float length = (isLTR ? baseFont : baseFontHebrew).GetWidthPoint(valueStr, (isLTR ? FontDefault : FontDefaultHebrew).CalculatedSize);
                                        if (relativeWidths[columnIndex] < length)
                                            relativeWidths[columnIndex] = length;
                                    }
                                }
                            }

                            // column names
                            List<string> columnNames = new List<string>(ExportHeaders);
                            float tableToHeaderRatio = 0.56F;
                            for (int i = 0; i < columnNames.Count; i++)
                            {
                                bool isLTR = IsLTR(columnNames[i]);
                                float columnNameWidth = (isLTR ? baseFont : baseFontHebrew).GetWidthPoint(columnNames[i], (isLTR ? FontDefault : FontDefaultHebrew).CalculatedSize);
                                if (relativeWidths[i] < columnNameWidth)
                                {
                                    if (columnNames[i].Length <= 10 && columnNames[i].Contains(' ') == false) // short header
                                    {
                                        relativeWidths[i] = columnNameWidth;
                                    }
                                    else if (relativeWidths[i] / columnNameWidth >= tableToHeaderRatio) // the header is not too long
                                    {
                                        relativeWidths[i] = columnNameWidth;
                                    }
                                    else if (columnNames[i].Contains(' ')) // the header is long but can be split down
                                    {
                                        string[] items = columnNames[i].Split(' ');

                                        if (items.Length == 2)
                                        {
                                            columnNames[i] = items[0] + "\n" + items[1];
                                            isLTR = IsLTR(items[0]);
                                            float length = (isLTR ? baseFont : baseFontHebrew).GetWidthPoint(items[0], (isLTR ? FontDefault : FontDefaultHebrew).CalculatedSize);
                                            if (relativeWidths[i] < length)
                                                relativeWidths[i] = length;
                                            isLTR = IsLTR(items[1]);
                                            length = (isLTR ? baseFont : baseFontHebrew).GetWidthPoint(items[1], (isLTR ? FontDefault : FontDefaultHebrew).CalculatedSize);
                                            if (relativeWidths[i] < length)
                                                relativeWidths[i] = length;
                                        }
                                        else if (items.Length == 3)
                                        {
                                            isLTR = IsLTR(items[0] + " " + items[1]);
                                            float length01 = (isLTR ? baseFont : baseFontHebrew).GetWidthPoint(items[0] + " " + items[1], (isLTR ? FontDefault : FontDefaultHebrew).CalculatedSize);

                                            isLTR = IsLTR(items[2]);
                                            float length2 = (isLTR ? baseFont : baseFontHebrew).GetWidthPoint(items[2], (isLTR ? FontDefault : FontDefaultHebrew).CalculatedSize);

                                            isLTR = IsLTR(items[0]);
                                            float length0 = (isLTR ? baseFont : baseFontHebrew).GetWidthPoint(items[0], (isLTR ? FontDefault : FontDefaultHebrew).CalculatedSize);

                                            isLTR = IsLTR(items[1] + " " + items[2]);
                                            float length12 = (isLTR ? baseFont : baseFontHebrew).GetWidthPoint(items[1] + " " + items[2], (isLTR ? FontDefault : FontDefaultHebrew).CalculatedSize);

                                            // the bigger ratio has more balance between the 2 text segments
                                            float ratio01 = length2 / length01;
                                            float ratio12 = length0 / length12;

                                            if (ratio01 > ratio12)
                                            {
                                                columnNames[i] = items[0] + " " + items[1] + "\n" + items[2];
                                                if (relativeWidths[i] < length01)
                                                    relativeWidths[i] = length01;
                                                if (relativeWidths[i] < length2)
                                                    relativeWidths[i] = length2;
                                            }
                                            else
                                            {
                                                columnNames[i] = items[0] + "\n" + items[1] + " " + items[2];
                                                if (relativeWidths[i] < length0)
                                                    relativeWidths[i] = length0;
                                                if (relativeWidths[i] < length12)
                                                    relativeWidths[i] = length12;
                                            }
                                        }
                                        else
                                        {
                                            for (int j = 0; j < items.Length; j++)
                                            {
                                                isLTR = IsLTR(items[j]);
                                                float length = (isLTR ? baseFont : baseFontHebrew).GetWidthPoint(items[j], (isLTR ? FontDefault : FontDefaultHebrew).CalculatedSize);

                                                if (relativeWidths[i] < length)
                                                {
                                                    if (items[j].Length <= 10) // short header
                                                    {
                                                        relativeWidths[i] = length;
                                                    }
                                                    else if (relativeWidths[i] / length >= tableToHeaderRatio) // the header is not too long
                                                    {
                                                        relativeWidths[i] = length;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            // balance columns
                            float thresholdRatioWidth = 0.08F;
                            float totalRelativeWidths = relativeWidths.Sum();
                            var ratioWidths = relativeWidths.Select((w, i) => new { Width = w, Ratio = w / totalRelativeWidths, Index = i });
                            if (ratioWidths.Any(x => x.Ratio < thresholdRatioWidth))
                            {
                                // (width1 + diff1) / (totalRelativeWidths + (count1 * diff1) + (count2 * diff2) + (count3 * diff3)) = minRatioWidth
                                // (width2 + diff2) / (totalRelativeWidths + (count1 * diff1) + (count2 * diff2) + (count3 * diff3)) = minRatioWidth
                                // (width3 + diff3) / (totalRelativeWidths + (count1 * diff1) + (count2 * diff2) + (count3 * diff3)) = minRatioWidth
                                // Ax = b
                                // A = [(count1 * minRatioWidth) - 1,  count2 * minRatioWidth,       count3 * minRatioWidth,
                                //       count1 * minRatioWidth,      (count2 * minRatioWidth) - 1,  count3 * minRatioWidth,
                                //       count1 * minRatioWidth,       count2 * minRatioWidth,      (count3 * minRatioWidth) - 1]
                                // x = [diff1, diff2, diff3]
                                // b = [width1 - (minRatioWidth * totalRelativeWidths), 
                                //      width2 - (minRatioWidth * totalRelativeWidths), 
                                //      width3 - (minRatioWidth * totalRelativeWidths)]

                                float minRatioWidth = ratioWidths.Where(x => x.Ratio < thresholdRatioWidth).Average(x => x.Ratio);
                                if (ratioWidths.Count(x => x.Ratio < minRatioWidth) > 0)
                                {
                                    var groups = ratioWidths.Where(x => x.Ratio < minRatioWidth).GroupBy(x => x.Width).Select(g => new { Width = g.Key, Count = g.Count() });
                                    int gcount = groups.Count();

                                    float[,] A = new float[gcount, gcount];
                                    for (int i = 0; i < gcount; i++)
                                    {
                                        int j = 0;
                                        foreach (var g in groups)
                                            A[i, j++] = minRatioWidth * g.Count;
                                        A[i, i] -= 1F;
                                    }

                                    float[] b = groups.Select(g => g.Width - (minRatioWidth * totalRelativeWidths)).ToArray();

                                    float[] diffs = Matrix<float>.Build.DenseOfArray(A).Solve(Vector<float>.Build.Dense(b)).ToArray();

                                    var items = groups.Zip(diffs, (g, diff) => new { Indices = ratioWidths.Where(x => x.Width == g.Width).Select(x => x.Index), diff }).Where(x => x.diff > 0);
                                    foreach (var item in items)
                                    {
                                        foreach (int index in item.Indices)
                                            relativeWidths[index] += item.diff;
                                    }
                                }
                            }

                            // page size
                            float cellPadding = 4F;
                            float pageWidth =
                                relativeWidths.Sum() +
                                (2 * cellPadding * ExportColumnsCount) + // cell padding
                                (60F * ExportColumnsCount);

                            // table WidthPercentage: 80F -> 100F
                            pageWidth = 0.8F * pageWidth;

                            int thresholdLinesA3Height = 53;
                            Rectangle pageSize = PageSize.A4;
                            if (pageWidth > PageSize.A4.Width)
                            {
                                int totalLinesCount = linesCount;
                                totalLinesCount += 1; // sum row

                                float pageHeight = (totalLinesCount <= thresholdLinesA3Height ? PageSize.A3.Height : (pageWidth <= PageSize.A1.Width ? PageSize.A2.Height : PageSize.A0.Height));
                                pageSize = new Rectangle(pageWidth, pageHeight);
                            }

                            // writer event
                            writer.PageEvent = new PdfPageEvent()
                            {
                                document = document,
                                now = now,
                                columnNames = columnNames,
                                PrintingFormat = PrintingFormat,
                                relativeWidths = relativeWidths,
                                cellPadding = cellPadding,
                                IsLTR = IsLTR,
                                linesCount = linesCount,
                                numerator = numerator,
                                FontDefault = FontDefault,
                                FontDefaultBold = FontDefaultBold,
                                FontOutline = FontOutline,
                                FontReportName = FontReportName,
                                FontSearchCriteria2 = FontSearchCriteria2,
                                FontDefaultHebrew = FontDefaultHebrew,
                                FontDefaultBoldHebrew = FontDefaultBoldHebrew,
                                FontOutlineHebrew = FontOutlineHebrew,
                                FontReportNameHebrew = FontReportNameHebrew,
                                FontSearchCriteria2Hebrew = FontSearchCriteria2Hebrew
                            };

                            // document
                            float pageMargin = 20F;
                            document.SetPageSize(pageSize);
                            document.SetMargins(pageMargin, pageMargin, pageMargin, pageMargin);
                            document.Open();

                            // meta information
                            if (string.IsNullOrEmpty(REPORT_NAME) == false)
                            {
                                document.AddTitle(REPORT_NAME);
                                document.AddSubject(REPORT_NAME);
                                document.AddAuthor(PDF_METADATA_AUTHOR);
                                document.AddCreator(PDF_METADATA_CREATOR);

                                if (string.IsNullOrEmpty(PDF_METADATA_KEYWORDS))
                                    document.AddKeywords(REPORT_NAME);
                                else
                                    document.AddKeywords(PDF_METADATA_KEYWORDS + "," + REPORT_NAME);
                            }

                            PdfPTable table = GetPDFTable(ExportColumnsCount);

                            // invoice rows
                            for (int rowIndex = 0; rowIndex < tableValues.Length && rowIndex < invoiceLinesCount; rowIndex++)
                            {
                                var row = tableValues[rowIndex];

                                for (int columnIndex = 0; columnIndex < row.Length; columnIndex++)
                                {
                                    string valueStr = row[columnIndex];

                                    int horizontalAlignment = (ExportHorizontalAlignmentRight[columnIndex] ? Rectangle.ALIGN_RIGHT : Rectangle.ALIGN_LEFT);

                                    bool isLTR = (horizontalAlignment == Rectangle.ALIGN_LEFT ? IsLTR(valueStr) : true);
                                    PdfPCell cell = new PdfPCell(new Phrase(valueStr, (isLTR ? FontDefault : FontDefaultHebrew)));
                                    cell.RunDirection = (isLTR ? PdfWriter.RUN_DIRECTION_LTR : PdfWriter.RUN_DIRECTION_RTL);
                                    cell.HorizontalAlignment = horizontalAlignment;
                                    cell.VerticalAlignment = Rectangle.ALIGN_MIDDLE;
                                    cell.Border = Rectangle.BOX;
                                    cell.NoWrap = false;
                                    cell.Padding = cellPadding;
                                    table.AddCell(cell);
                                }

                                var line = invoiceLines.ElementAt(rowIndex);
                                bool isManualCalculation = line.Is_Manual_Calculation == 1;
                                if (isManualCalculation)
                                {
                                    PdfPCell cell = table.Rows[rowIndex].GetCells()[13]; // deduction
                                    if (cell.Phrase != null && cell.Phrase.Chunks != null && cell.Phrase.Chunks.Count > 0)
                                        cell.Phrase.Chunks[0].Font = (cell.RunDirection == PdfWriter.RUN_DIRECTION_LTR ? FontDefaultBold : FontDefaultBoldHebrew);
                                }
                            }

                            // mileage rows
                            for (int rowIndex = invoiceLinesCount; rowIndex < tableValues.Length - 1; rowIndex++)
                            {
                                var row = tableValues[rowIndex];

                                for (int columnIndex = 0; columnIndex < row.Length; columnIndex++)
                                {
                                    string valueStr = row[columnIndex];

                                    int horizontalAlignment = (ExportHorizontalAlignmentRight[columnIndex] ? Rectangle.ALIGN_RIGHT : Rectangle.ALIGN_LEFT);

                                    bool isLTR = (horizontalAlignment == Rectangle.ALIGN_LEFT ? IsLTR(valueStr) : true);
                                    PdfPCell cell = new PdfPCell(new Phrase(valueStr, (isLTR ? FontDefault : FontDefaultHebrew)));
                                    cell.RunDirection = (isLTR ? PdfWriter.RUN_DIRECTION_LTR : PdfWriter.RUN_DIRECTION_RTL);
                                    cell.HorizontalAlignment = horizontalAlignment;
                                    cell.VerticalAlignment = Rectangle.ALIGN_MIDDLE;
                                    cell.Border = Rectangle.BOX;
                                    cell.NoWrap = false;
                                    cell.Padding = cellPadding;
                                    table.AddCell(cell);
                                }

                                PdfPCell[] cells = table.Rows[rowIndex].GetCells();
                                foreach (PdfPCell cell in cells)
                                {
                                    if (cell.Phrase != null && cell.Phrase.Chunks != null && cell.Phrase.Chunks.Count > 0)
                                        cell.Phrase.Chunks[0].Font = (cell.RunDirection == PdfWriter.RUN_DIRECTION_LTR ? FontDefaultBoldGreen : FontDefaultBoldGreenHebrew);
                                }
                            }

                            // sum row
                            var sumRow = tableValues[tableValues.Length - 1];

                            for (int columnIndex = 0; columnIndex < sumRow.Length; columnIndex++)
                            {
                                string valueStr = sumRow[columnIndex];

                                bool isLTR = IsLTR(valueStr);
                                PdfPCell cell = new PdfPCell(new Phrase(valueStr, (isLTR ? FontDefault : FontDefaultHebrew)));
                                cell.RunDirection = (isLTR ? PdfWriter.RUN_DIRECTION_LTR : PdfWriter.RUN_DIRECTION_RTL);
                                cell.HorizontalAlignment = Rectangle.ALIGN_RIGHT;
                                cell.VerticalAlignment = Rectangle.ALIGN_MIDDLE;
                                cell.Border = Rectangle.BOX;
                                cell.BackgroundColor = SumRowColor;
                                cell.NoWrap = false;
                                cell.Padding = cellPadding;
                                table.AddCell(cell);
                            }

                            table.SetWidths(relativeWidths);

                            document.Add(table);

                            document.Close();
                            writer.Close();
                        }
                    }

                    return ms.ToArray();
                }
            }

            public static bool IsLTR(string value)
            {
                if (string.IsNullOrEmpty(value))
                    return true;

                if (value.Any(c => 'א' <= c && c <= 'ת'))
                    return false;

                return true;
            }
        }

        #endregion

        #region Excel

        public static class ExcelExportHelper
        {
            private static readonly System.Drawing.Color HeaderColor = System.Drawing.Color.FromArgb(112, 173, 71);
            private static readonly System.Drawing.Color SearchCriteriaColor = System.Drawing.Color.FromArgb(31, 112, 75);
            private static readonly System.Drawing.Color OutlineHeaderColor = System.Drawing.Color.FromArgb(198, 224, 180);
            private static readonly System.Drawing.Color OutlineHeaderFontColor = System.Drawing.Color.FromArgb(63, 63, 63);
            private static readonly System.Drawing.Color OutlineColor = System.Drawing.Color.FromArgb(245, 245, 245);
            private static readonly System.Drawing.Color SumRowColor = System.Drawing.Color.FromArgb(143, 209, 49);
            private static readonly System.Drawing.Color Green = System.Drawing.Color.FromArgb(51, 153, 102);

            private const string TextFormat = "@";
            private const string ShortDateFormat = "dd/mm/yyyy";
            private const string LongDateFormat = "dd/mm/yyyy hh:mm:ss";
            private const string PrintingFormat = "dd/mm/yyyy hh:mm";
            private const string IntFormat = "#,##0";
            private const string DoubleFormat = "#,##0.000000";
            private const string CurrencyIntFormat = "[$£-809]#,##0;[$£-809]-#,##0";
            private const string CurrencyDoubleFormat = "[$£-809]#,##0.000000;[$£-809]-#,##0.000000";
            private const string PercentageIntFormat = "0%";
            private const string PercentageDoubleFormat = "0.000000%";

            public static byte[] GetExcel(int numerator, string customerName, IEnumerable<InvoiceLine> invoiceLines, IEnumerable<FSCMileageLine> mileageLines, DateTime now)
            {
                using (var wb = new ExcelPackage())
                {
                    var ws = wb.Workbook.Worksheets.Add(REPORT_NAME);
                    ws.View.ShowGridLines = false;

                    int rowIndex = 2;

                    // printing date
                    var printingDateRow = ws.Cells[rowIndex, 1, rowIndex, ExportColumnsCount];
                    printingDateRow.Merge = true;
                    printingDateRow.Value = now;
                    printingDateRow.Style.Numberformat.Format = PrintingFormat;
                    printingDateRow.Style.Font.Size = 10;
                    printingDateRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    printingDateRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    printingDateRow.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    printingDateRow.Style.WrapText = false;

                    rowIndex++;

                    // report name
                    var reportNameRow = ws.Cells[rowIndex, 1, rowIndex, ExportColumnsCount];
                    reportNameRow.Merge = true;
                    reportNameRow.Value = REPORT_NAME;
                    reportNameRow.Style.Numberformat.Format = TextFormat;
                    reportNameRow.Style.Font.Size = 13.5F;
                    reportNameRow.Style.Font.Bold = true;
                    reportNameRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    reportNameRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    reportNameRow.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    reportNameRow.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    reportNameRow.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    reportNameRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    reportNameRow.Style.Fill.BackgroundColor.SetColor(HeaderColor);
                    reportNameRow.Style.WrapText = false;

                    rowIndex++;

                    // search criteria
                    var searchCriteriaRow = ws.Cells[rowIndex, 1, rowIndex, ExportColumnsCount];
                    searchCriteriaRow.Merge = true;
                    searchCriteriaRow.Style.Numberformat.Format = TextFormat;
                    searchCriteriaRow.Style.Font.Size = 10;
                    searchCriteriaRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    searchCriteriaRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    searchCriteriaRow.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    searchCriteriaRow.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    searchCriteriaRow.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    searchCriteriaRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    searchCriteriaRow.Style.Fill.BackgroundColor.SetColor(HeaderColor);
                    searchCriteriaRow.Style.WrapText = true;
                    ws.Row(rowIndex).Height *= 2;
                    searchCriteriaRow.IsRichText = true;

                    int invoiceLinesCount = invoiceLines.Count();
                    int mileageLinesCount = mileageLines.Count();
                    int linesCount = invoiceLinesCount + mileageLinesCount;
                    searchCriteriaRow.RichText.Add(string.Format(SEARCH_CRITERIA_1, numerator));
                    searchCriteriaRow.RichText.Add(string.Format(SEARCH_CRITERIA_2, linesCount)).Color = SearchCriteriaColor;

                    rowIndex++;

                    // headers
                    for (int columnIndex = 1; columnIndex <= ExportColumnsCount; columnIndex++)
                    {
                        string header = ExportHeaders[columnIndex - 1];
                        var cell = ws.Cells[rowIndex, columnIndex];
                        cell.Value = header;
                        cell.Style.Numberformat.Format = TextFormat;
                        cell.Style.Font.Size = 10;
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(HeaderColor);
                        cell.Style.WrapText = false;
                    }

                    rowIndex++;

                    // invoice rows
                    foreach (var line in invoiceLines)
                    {
                        ws.Cells[rowIndex, 1].Value = line.Customer_Name;
                        ws.Cells[rowIndex, 2].Value = line.Site_Name;
                        ws.Cells[rowIndex, 3].Value = line.Value_Date;
                        ws.Cells[rowIndex, 4].Value = line.Delivery_Date;
                        ws.Cells[rowIndex, 5].Value = line.Numerator;
                        ws.Cells[rowIndex, 6].Value = line.Order_Numerator;
                        ws.Cells[rowIndex, 7].Value = line.Haulier_Name;
                        ws.Cells[rowIndex, 8].Value = line.Product_Name;
                        ws.Cells[rowIndex, 9].Value = line.Quantity;
                        ws.Cells[rowIndex, 10].Value = line.Price;
                        ws.Cells[rowIndex, 11].Value = line.Total;
                        ws.Cells[rowIndex, 12].Value = line.Internal_Use_Desc;
                        ws.Cells[rowIndex, 13].Value = line.Tax_Rates_Percent;
                        ws.Cells[rowIndex, 14].Value = line.Deduction;
                        ws.Cells[rowIndex, 15].Value = line.Final_Sales_Tax;
                        ws.Cells[rowIndex, 16].Value = line.Sales_Tax_Amount;
                        ws.Cells[rowIndex, 17].Value = line.Comments;
                        ws.Cells[rowIndex, 18].Value = line.Tax_Exemption_PTE;
                        ws.Cells[rowIndex, 19].Value = line.From_Site_Name;
                        ws.Cells[rowIndex, 20].Value = line.Sales_Tax_State;
                        ws.Cells[rowIndex, 21].Value = line.Sales_Tax_County;
                        ws.Cells[rowIndex, 22].Value = line.Sales_Tax_City;

                        for (int columnIndex = 1; columnIndex <= ExportColumnsCount; columnIndex++)
                        {
                            var cell = ws.Cells[rowIndex, columnIndex];
                            cell.Style.Numberformat.Format = TextFormat;
                            cell.Style.Font.Size = 10;
                            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            cell.Style.WrapText = false;

                            if (ExportHorizontalAlignmentRight[columnIndex - 1])
                                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            else
                                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }

                        // Value_Date
                        using (var cell = ws.Cells[rowIndex, 3])
                            cell.Style.Numberformat.Format = ShortDateFormat;

                        // Delivery_Date
                        using (var cell = ws.Cells[rowIndex, 4])
                            cell.Style.Numberformat.Format = ShortDateFormat;

                        // Quantity
                        using (var cell = ws.Cells[rowIndex, 9])
                            cell.Style.Numberformat.Format = IntFormat;

                        // Price
                        using (var cell = ws.Cells[rowIndex, 10])
                            cell.Style.Numberformat.Format = CurrencyDoubleFormat;

                        // Total
                        using (var cell = ws.Cells[rowIndex, 11])
                            cell.Style.Numberformat.Format = CurrencyDoubleFormat;

                        // Tax_Rates_Percent
                        using (var cell = ws.Cells[rowIndex, 13])
                            cell.Style.Numberformat.Format = PercentageDoubleFormat;

                        // Deduction
                        using (var cell = ws.Cells[rowIndex, 14])
                            cell.Style.Numberformat.Format = PercentageDoubleFormat;

                        // Final_Sales_Tax
                        using (var cell = ws.Cells[rowIndex, 15])
                            cell.Style.Numberformat.Format = PercentageDoubleFormat;

                        // Sales_Tax_Amount
                        using (var cell = ws.Cells[rowIndex, 16])
                            cell.Style.Numberformat.Format = CurrencyDoubleFormat;

                        // Sales_Tax_State
                        using (var cell = ws.Cells[rowIndex, 20])
                            cell.Style.Numberformat.Format = DoubleFormat;

                        // Sales_Tax_County
                        using (var cell = ws.Cells[rowIndex, 21])
                            cell.Style.Numberformat.Format = DoubleFormat;

                        // Sales_Tax_City
                        using (var cell = ws.Cells[rowIndex, 22])
                            cell.Style.Numberformat.Format = DoubleFormat;

                        bool isManualCalculation = line.Is_Manual_Calculation == 1;
                        if (isManualCalculation)
                            ws.Cells[rowIndex, 14].Style.Font.Bold = true; // deduction

                        rowIndex++;
                    }

                    // mileage rows
                    foreach (var line in mileageLines)
                    {
                        ws.Cells[rowIndex, 1].Value = customerName;
                        ws.Cells[rowIndex, 5].Value = line.Numerator;
                        ws.Cells[rowIndex, 8].Value = "Fuel Surcharge";
                        ws.Cells[rowIndex, 9].Value = line.Priority_FSC_Mileage;
                        ws.Cells[rowIndex, 10].Value = line.Priority_FSC_Price;
                        ws.Cells[rowIndex, 11].Value = line.Priority_FSC_Price;

                        for (int columnIndex = 1; columnIndex <= ExportColumnsCount; columnIndex++)
                        {
                            var cell = ws.Cells[rowIndex, columnIndex];
                            cell.Style.Numberformat.Format = TextFormat;
                            cell.Style.Font.Size = 10;
                            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                            cell.Style.WrapText = false;

                            if (ExportHorizontalAlignmentRight[columnIndex - 1])
                                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                            else
                                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        }

                        // Quantity
                        using (var cell = ws.Cells[rowIndex, 9])
                            cell.Style.Numberformat.Format = IntFormat;

                        // Price
                        using (var cell = ws.Cells[rowIndex, 10])
                            cell.Style.Numberformat.Format = CurrencyDoubleFormat;

                        // Total
                        using (var cell = ws.Cells[rowIndex, 11])
                            cell.Style.Numberformat.Format = CurrencyDoubleFormat;

                        ws.Cells[rowIndex, 1, rowIndex, ws.Dimension.End.Column].Style.Font.Color.SetColor(Green);
                        ws.Cells[rowIndex, 1, rowIndex, ExportColumnsCount].Style.Font.Bold = true;

                        rowIndex++;
                    }

                    // sum row
                    int invoiceLinesSumQuantity = invoiceLines.Sum(x => x.Quantity) ?? 0;
                    int mileageLinesSumQuantity = mileageLines.Sum(x => x.Priority_FSC_Mileage);
                    int sumQuantity = invoiceLinesSumQuantity + mileageLinesSumQuantity;

                    decimal invoiceLinesSumTotal = invoiceLines.Sum(x => x.Total) ?? 0;
                    decimal mileageLinesSumTotal = mileageLines.Sum(x => x.Priority_FSC_Price);
                    decimal sumTotal = invoiceLinesSumTotal + mileageLinesSumTotal;

                    for (int columnIndex = 1; columnIndex <= ExportColumnsCount; columnIndex++)
                    {
                        var cell = ws.Cells[rowIndex, columnIndex];
                        cell.Style.Font.Size = 10;
                        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(SumRowColor);
                        cell.Style.WrapText = false;
                    }

                    // Quantity
                    using (var cell = ws.Cells[rowIndex, 9])
                    {
                        cell.Value = sumQuantity;
                        cell.Style.Numberformat.Format = IntFormat;
                    }

                    // Total
                    using (var cell = ws.Cells[rowIndex, 11])
                    {
                        cell.Value = sumTotal;
                        cell.Style.Numberformat.Format = CurrencyDoubleFormat;
                    }

                    // auto fit
                    for (int columnIndex = 1; columnIndex <= ExportColumnsCount; columnIndex++)
                        ws.Column(columnIndex).AutoFit();

                    return wb.GetAsByteArray();
                }
            }
        }

        #endregion

        #region File

        public static class FileExportHelper
        {
            // http://blogs.iis.net/robert_mcmurray/bad-characters-to-use-in-web-based-filenames
            private static readonly Regex reservedChars = new Regex(string.Format("[{0}]", Regex.Escape(";/?:@&=+$,")), RegexOptions.Compiled);
            private static readonly Regex excludedChars = new Regex(string.Format("[{0}]", Regex.Escape("][<>#%\"{}|\\^'")), RegexOptions.Compiled);
            private static readonly Regex invalidFileNameChars = new Regex(string.Format("[{0}]", Regex.Escape(new string(System.IO.Path.GetInvalidFileNameChars()))), RegexOptions.Compiled);

            public static string CleanFileName(string fileName)
            {
                if (string.IsNullOrEmpty(fileName))
                    return System.IO.Path.GetFileNameWithoutExtension(System.IO.Path.GetRandomFileName());
                fileName = ReplaceAccentedChars(fileName);
                fileName = reservedChars.Replace(fileName, "_");
                fileName = excludedChars.Replace(fileName, "_");
                fileName = invalidFileNameChars.Replace(fileName, "_");
                return fileName;
            }

            #region Accented Chars

            // http://tlt.its.psu.edu/suggestions/international/accents/codealt.html#accent

            private static readonly Regex regex_A = new Regex("[ÀÁÂÃÄÅ]", RegexOptions.Compiled);
            private static readonly Regex regex_E = new Regex("[ÈÉÊË]", RegexOptions.Compiled);
            private static readonly Regex regex_I = new Regex("[ÌÍÎÏ]", RegexOptions.Compiled);
            private static readonly Regex regex_O = new Regex("[ÒÓÔÕÖØ]", RegexOptions.Compiled);
            private static readonly Regex regex_U = new Regex("[ÙÚÛÜ]", RegexOptions.Compiled);
            private static readonly Regex regex_Y = new Regex("[ÝŸ]", RegexOptions.Compiled);
            private static readonly Regex regex_N = new Regex("[Ñ]", RegexOptions.Compiled);
            private static readonly Regex regex_a = new Regex("[àáâãäªå]", RegexOptions.Compiled);
            private static readonly Regex regex_e = new Regex("[èéêë]", RegexOptions.Compiled);
            private static readonly Regex regex_i = new Regex("[ìíîï]", RegexOptions.Compiled);
            private static readonly Regex regex_o = new Regex("[òóôõöºø]", RegexOptions.Compiled);
            private static readonly Regex regex_u = new Regex("[ùúûü]", RegexOptions.Compiled);
            private static readonly Regex regex_y = new Regex("[ýÿ]", RegexOptions.Compiled);
            private static readonly Regex regex_n = new Regex("[ñ]", RegexOptions.Compiled);
            private static readonly Regex regex_C = new Regex("[Ç]", RegexOptions.Compiled);
            private static readonly Regex regex_c = new Regex("[ç]", RegexOptions.Compiled);
            private static readonly Regex regex_OE = new Regex("[Œ]", RegexOptions.Compiled);
            private static readonly Regex regex_oe = new Regex("[œ]", RegexOptions.Compiled);
            private static readonly Regex regex_ss = new Regex("[ß]", RegexOptions.Compiled);
            private static readonly Regex regex_AE = new Regex("[Æ]", RegexOptions.Compiled);
            private static readonly Regex regex_ae = new Regex("[æ]", RegexOptions.Compiled);
            private static readonly Regex regex_TH = new Regex("[Þ]", RegexOptions.Compiled);
            private static readonly Regex regex_th = new Regex("[þ]", RegexOptions.Compiled);
            private static readonly Regex regex_D = new Regex("[Ð]", RegexOptions.Compiled);
            private static readonly Regex regex_d = new Regex("[ð]", RegexOptions.Compiled);
            private static readonly Regex regex_S = new Regex("[Š]", RegexOptions.Compiled);
            private static readonly Regex regex_s = new Regex("[š]", RegexOptions.Compiled);
            private static readonly Regex regex_Z = new Regex("[Ž]", RegexOptions.Compiled);
            private static readonly Regex regex_z = new Regex("[ž]", RegexOptions.Compiled);

            private static string ReplaceAccentedChars(string str)
            {
                str = regex_A.Replace(str, "A");
                str = regex_E.Replace(str, "E");
                str = regex_I.Replace(str, "I");
                str = regex_O.Replace(str, "O");
                str = regex_U.Replace(str, "U");
                str = regex_Y.Replace(str, "Y");
                str = regex_N.Replace(str, "N");
                str = regex_a.Replace(str, "a");
                str = regex_e.Replace(str, "e");
                str = regex_i.Replace(str, "i");
                str = regex_o.Replace(str, "o");
                str = regex_u.Replace(str, "u");
                str = regex_y.Replace(str, "y");
                str = regex_n.Replace(str, "n");
                str = regex_C.Replace(str, "C");
                str = regex_c.Replace(str, "c");
                str = regex_OE.Replace(str, "OE");
                str = regex_oe.Replace(str, "oe");
                str = regex_ss.Replace(str, "ss");
                str = regex_AE.Replace(str, "AE");
                str = regex_ae.Replace(str, "ae");
                str = regex_TH.Replace(str, "TH");
                str = regex_th.Replace(str, "th");
                str = regex_D.Replace(str, "D");
                str = regex_d.Replace(str, "d");
                str = regex_S.Replace(str, "S");
                str = regex_s.Replace(str, "s");
                str = regex_Z.Replace(str, "Z");
                str = regex_z.Replace(str, "z");
                return str;
            }

            #endregion
        }

        #endregion
    }
}
