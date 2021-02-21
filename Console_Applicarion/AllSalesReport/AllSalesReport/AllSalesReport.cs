using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using Dapper;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

namespace AllSalesReport
{
    class AllSalesReport
    {
        #region Run

        public void Run(string[] args)
        {
            Console.WriteLine("Version - " + Assembly.GetExecutingAssembly().GetName().Version.ToString());
            DateTime now = DateTime.Now;

            try
            {
                string exportPath = ConfigurationManager.AppSettings["Export_Path"];
                bool exportEnabled = (string.IsNullOrEmpty(exportPath) == false);

                bool mailEnabled = false;
                mailEnabled = bool.TryParse(ConfigurationManager.AppSettings["Mail_Enabled"], out mailEnabled) && mailEnabled;

                if (exportEnabled == false && mailEnabled == false)
                    return;

                var data = GetData();

                var details = data.Item1;
                var eastDetails = data.Item2;
                var westDetails = data.Item3;
                var customerSummary = data.Item4;
                var eastCustomerSummary = data.Item5;
                var westCustomerSummary = data.Item6;
                var customerSummaryLight = data.Item7;
                var lyVsTYBySize = data.Rest.Item1;
                var eastLyVsTYBySize = data.Rest.Item2;
                var westLyVsTYBySize = data.Rest.Item3;
                var relCustomerCategory = data.Rest.Item4;
                var salesReps = data.Rest.Item5;

                var customersOpenPool = relCustomerCategory.Where(x => x.Category_Code == (int)Categories.Open_Pool);
                var customersDedicatedLoops = relCustomerCategory.Where(x => x.Category_Code == (int)Categories.Dedicated_Loops);
                var customersInsideSales = relCustomerCategory.Where(x => x.Category_Code == (int)Categories.Inside_Sales);
                var customersClosedLoops = relCustomerCategory.Where(x => x.Category_Code == (int)Categories.Closed_Loops);
                var customersMeat = relCustomerCategory.Where(x => x.Category_Code == (int)Categories.Meat);
                var customersOthers = relCustomerCategory.Where(x => x.Category_Code == (int)Categories.Others);

                var eastSalesReps = salesReps.Where(x => x.Is_East == 1);
                var westSalesReps = salesReps.Where(x => x.Is_West == 1);

                var model = new Model()
                {
                    Details = details,
                    EastDetails = eastDetails,
                    WestDetails = westDetails,
                    CustomerSummary = customerSummary,
                    EastCustomerSummary = eastCustomerSummary,
                    WestCustomerSummary = westCustomerSummary,
                    CustomerSummaryLight = customerSummaryLight,
                    LyVsTYBySize = lyVsTYBySize,
                    EastLyVsTYBySize = eastLyVsTYBySize,
                    WestLyVsTYBySize = westLyVsTYBySize,
                    RelCustomerCategory = relCustomerCategory,
                    SalesReps = salesReps,
                    CustomersOpenPool = customersOpenPool,
                    CustomersDedicatedLoops = customersDedicatedLoops,
                    CustomersInsideSales = customersInsideSales,
                    CustomersClosedLoops = customersClosedLoops,
                    CustomersMeat = customersMeat,
                    CustomersOthers = customersOthers,
                    EastSalesReps = eastSalesReps,
                    WestSalesReps = westSalesReps
                };

                bool isLight =
                    args != null &&
                    args.Any(a => string.Compare(a, "Light", true) == 0);

                if (isLight)
                {
                    HandleExcel(LightReport, model, "ALL Sales Light Report", now, exportEnabled, exportPath, mailEnabled, "Mail_Light_Report");
                }
                else
                {
                    HandleExcel(ALLSalesReport, model, "ALL Sales Report", now, exportEnabled, exportPath, mailEnabled, "Mail_To_Address");
                    HandleExcel(TonyMosco, model, "Tony Mosco", now, exportEnabled, exportPath, mailEnabled, "Mail_Tony_Mosco");
                    /*HandleExcel(DanaEmbry, model, "Dana Embry", now, exportEnabled, exportPath, mailEnabled, "Mail_Dana_Embry");
                    HandleExcel(BrandonOsteen, model, "Brandon O'steen", now, exportEnabled, exportPath, mailEnabled, "Mail_Brandon_Osteen");*/
                    HandleExcel(TravisWhite, model, "Travis White", now, exportEnabled, exportPath, mailEnabled, "Mail_Travis_White");
                    HandleExcel(FrankNiolet, model, "Frank Niolet", now, exportEnabled, exportPath, mailEnabled, "Mail_Frank_Niolet");
                    HandleExcel(SarahGonzalez, model, "Sarah Gonzalez", now, exportEnabled, exportPath, mailEnabled, "Mail_Sarah_Gonzalez");
                    HandleExcel(GabrielGrijalva, model, "Gabriel Grijalva", now, exportEnabled, exportPath, mailEnabled, "Mail_Gabriel_Grijalva");
                    HandleExcel(PaulPederson, model, "Paul Pederson", now, exportEnabled, exportPath, mailEnabled, "Mail_Paul_Pederson");
                    HandleExcel(ConnieCeniceros, model, "Connie Ceniceros", now, exportEnabled, exportPath, mailEnabled, "Mail_Connie_Ceniceros");
                    HandleExcel(LisaDennis, model, "Lisa Dennis", now, exportEnabled, exportPath, mailEnabled, "Mail_Lisa_Dennis");
                    HandleExcel(EricBiddiscombe, model, "Eric Biddiscombe", now, exportEnabled, exportPath, mailEnabled, "Mail_Eric_Biddiscombe");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(string.Format("{0:yyyy-MM-dd HH:mm:ss} - {1}", now, ex.Message));
            }
        }

        private void HandleExcel(Func<Model, byte[]> GetExcel, Model model, string excelName, DateTime now, bool exportEnabled, string exportPath, bool mailEnabled, string mailToAddressKey)
        {
            byte[] excel = GetExcel(model);

            string name = string.Format("{0} {1:MMddyyyy}", excelName, now);
            string fileName = name + ".xlsx";

            if (exportEnabled)
            {
                string filePath = exportPath.TrimEnd('\\') + "\\" + fileName;
                ExportExcel(excel, fileName, now, filePath);
            }

            if (mailEnabled)
            {
                string mailToAddress = ConfigurationManager.AppSettings[mailToAddressKey];
                MailExcel(excel, fileName, now, name, mailToAddress);
            }
        }

        private void ExportExcel(byte[] excel, string fileName, DateTime now, string filePath)
        {
            try
            {
                File.WriteAllBytes(filePath, excel);
                Console.WriteLine(string.Format("{0:yyyy-MM-dd HH:mm:ss} - {1}: {2}", now, "Export Succeeded", fileName));
            }
            catch (Exception exExport)
            {
                Console.WriteLine(string.Format("{0:yyyy-MM-dd HH:mm:ss} - {1}: {2}", now, "Export Failed", exExport.Message));
            }
        }

        private void MailExcel(byte[] excel, string fileName, DateTime now, string name, string mailToAddress)
        {
            try
            {
                using (MemoryStream contentStream = new MemoryStream())
                {
                    contentStream.Write(excel, 0, excel.Length);
                    contentStream.Position = 0;

                    string mediaType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    MailSender.SendMail(name, string.Empty, mailToAddress, contentStream, fileName, mediaType);
                    Console.WriteLine(string.Format("{0:yyyy-MM-dd HH:mm:ss} - {1}: {2}", now, "Mail Succeeded", fileName));
                }
            }
            catch (Exception exMail)
            {
                Console.WriteLine(string.Format("{0:yyyy-MM-dd HH:mm:ss} - {1}: {2}", now, "Mail Failed", exMail.Message));
            }
        }

        #endregion

        #region Data

        private
            Tuple<IEnumerable<DetailsLine>, IEnumerable<DetailsLine>, IEnumerable<DetailsLine>,
            IEnumerable<CustomerSummaryLine>, IEnumerable<CustomerSummaryLine>, IEnumerable<CustomerSummaryLine>,
            IEnumerable<CustomerSummaryLineLight>,
            Tuple<IEnumerable<LYVsTYBySizeLine>, IEnumerable<LYVsTYBySizeLine>, IEnumerable<LYVsTYBySizeLine>,
            IEnumerable<RelCustomerCategory>, IEnumerable<SalesRep>>> GetData()
        {
            var data = SqlHelper.QueryMultipleSP<
                DetailsLine, DetailsLine, DetailsLine,
                CustomerSummaryLine, CustomerSummaryLine, CustomerSummaryLine,
                CustomerSummaryLineLight,
                LYVsTYBySizeLine, LYVsTYBySizeLine, LYVsTYBySizeLine,
                RelCustomerCategory, SalesRep>("sp_all_sales_report");

            TrimDetailsLines(data.Item1);
            TrimDetailsLines(data.Item2);
            TrimDetailsLines(data.Item3);

            TrimCustomerSummaryLines(data.Item4);
            TrimCustomerSummaryLines(data.Item5);
            TrimCustomerSummaryLines(data.Item6);

            TrimCustomerSummaryLinesLight(data.Item7);

            TrimLYVsTYBySizeLines(data.Rest.Item1);
            TrimLYVsTYBySizeLines(data.Rest.Item2);
            TrimLYVsTYBySizeLines(data.Rest.Item3);

            TrimRelCustomerCategoryLines(data.Rest.Item4);

            TrimSalesRepLines(data.Rest.Item5);

            return data;
        }

        private void TrimDetailsLines(IEnumerable<DetailsLine> lines)
        {
            foreach (var line in lines)
            {
                line.Customer_Name = line.Customer_Name.Trim();
                line.Office_Sale_Desc = line.Office_Sale_Desc.Trim();
                line.Product_Name = line.Product_Name.Trim();
            }
        }

        private void TrimCustomerSummaryLines(IEnumerable<CustomerSummaryLine> lines)
        {
            foreach (var line in lines)
            {
                line.Customer_Name = line.Customer_Name.Trim();
                line.Office_Sale_Desc = line.Office_Sale_Desc.Trim();
            }
        }

        private void TrimCustomerSummaryLinesLight(IEnumerable<CustomerSummaryLineLight> lines)
        {
            foreach (var line in lines)
            {
                line.Customer_Name = line.Customer_Name.Trim();
                line.Office_Sale_Desc = line.Office_Sale_Desc.Trim();
                line.Loop_Desc = line.Loop_Desc.Trim();
                line.Commodity_Desc = line.Commodity_Desc.Trim();
            }
        }

        private void TrimLYVsTYBySizeLines(IEnumerable<LYVsTYBySizeLine> lines)
        {
            foreach (var line in lines)
            {
                line.Product_Name = line.Product_Name.Trim();
            }
        }

        private void TrimRelCustomerCategoryLines(IEnumerable<RelCustomerCategory> lines)
        {
            foreach (var line in lines)
            {
                line.Customer_Name = line.Customer_Name.Trim();
                line.Category_Desc = line.Category_Desc.Trim();
            }
        }

        private void TrimSalesRepLines(IEnumerable<SalesRep> lines)
        {
            foreach (var line in lines)
            {
                line.Office_Sale_Desc = line.Office_Sale_Desc.Trim();
            }
        }

        #endregion

        #region Reports

        private byte[] ALLSalesReport(Model model)
        {
            using (ExcelPackage wb = new ExcelPackage())
            {
                wb.Compression = CompressionLevel.BestCompression;

                EastOpenPool(model, wb);
                TonyOpenPool(model, wb);
                /*DanaOpenPool(model, wb);
                BrandonOpenPool(model, wb);*/
                FrankOpenPool(model, wb);
                /*OpenTerritoryOpenPool(model, wb);*/
                EastLyVsTYBySize(model, wb);
                DedicatedLoops(model, wb);
                Travis(model, wb);
                East(model, wb);
                WestOpenPool(model, wb);
                SarahOpenPool(model, wb);
                GabeOpenPool(model, wb);
                ConnieOpenPool(model, wb);
                LisaOpenPool(model, wb);
                WestLyVsTYBySize(model, wb);
                ClosedLoops(model, wb);
                West(model, wb);
                Meat(model, wb);
                Eggs(model, wb);
                Others(model, wb);
                Eric(model, wb);
                AllOpenPool(model, wb);
                AllLyVsTYBySize(model, wb);
                return wb.GetAsByteArray();
            }
        }

        private byte[] TonyMosco(Model model)
        {
            using (ExcelPackage wb = new ExcelPackage())
            {
                wb.Compression = CompressionLevel.BestCompression;

                EastOpenPool(model, wb);
                TonyOpenPool(model, wb);
                /*BrandonOpenPool(model, wb);
                OpenTerritoryOpenPool(model, wb);*/
                EastLyVsTYBySize(model, wb);
                DedicatedLoops(model, wb);
                Travis(model, wb);
                East(model, wb);
                return wb.GetAsByteArray();
            }
        }

        /*private byte[] DanaEmbry(Model model)
        {
            using (ExcelPackage wb = new ExcelPackage())
            {
                wb.Compression = CompressionLevel.BestCompression;

                DanaOpenPool(model, wb);
                Others(model, wb);
                return wb.GetAsByteArray();
            }
        }

        private byte[] BrandonOsteen(Model model)
        {
            using (ExcelPackage wb = new ExcelPackage())
            {
                wb.Compression = CompressionLevel.BestCompression;

                BrandonOpenPool(model, wb);
                return wb.GetAsByteArray();
            }
        }*/

        private byte[] TravisWhite(Model model)
        {
            using (ExcelPackage wb = new ExcelPackage())
            {
                wb.Compression = CompressionLevel.BestCompression;

                Travis(model, wb);
                return wb.GetAsByteArray();
            }
        }

        private byte[] FrankNiolet(Model model)
        {
            using (ExcelPackage wb = new ExcelPackage())
            {
                wb.Compression = CompressionLevel.BestCompression;

                DedicatedLoops(model, wb);
                FrankOpenPool(model, wb);
                Eggs(model, wb);
                return wb.GetAsByteArray();
            }
        }

        private byte[] SarahGonzalez(Model model)
        {
            using (ExcelPackage wb = new ExcelPackage())
            {
                wb.Compression = CompressionLevel.BestCompression;

                EastOpenPool(model, wb);
                DedicatedLoops(model, wb);
                Travis(model, wb);
                WestOpenPool(model, wb);
                SarahOpenPool(model, wb);
                GabeOpenPool(model, wb);
                WestLyVsTYBySize(model, wb);
                ClosedLoops(model, wb);
                West(model, wb);
                Others(model, wb);
                AllOpenPool(model, wb);
                return wb.GetAsByteArray();
            }
        }

        private byte[] GabrielGrijalva(Model model)
        {
            using (ExcelPackage wb = new ExcelPackage())
            {
                wb.Compression = CompressionLevel.BestCompression;

                GabeOpenPool(model, wb);
                return wb.GetAsByteArray();
            }
        }

        private byte[] PaulPederson(Model model)
        {
            using (ExcelPackage wb = new ExcelPackage())
            {
                wb.Compression = CompressionLevel.BestCompression;

                Meat(model, wb);
                return wb.GetAsByteArray();
            }
        }

        private byte[] ConnieCeniceros(Model model)
        {
            using (ExcelPackage wb = new ExcelPackage())
            {
                wb.Compression = CompressionLevel.BestCompression;

                DedicatedLoops(model, wb);
                WestOpenPool(model, wb);
                ConnieOpenPool(model, wb);
                ClosedLoops(model, wb);
                West(model, wb);
                AllOpenPool(model, wb);
                return wb.GetAsByteArray();
            }
        }

        private byte[] LisaDennis(Model model)
        {
            using (ExcelPackage wb = new ExcelPackage())
            {
                wb.Compression = CompressionLevel.BestCompression;

                DedicatedLoops(model, wb);
                WestOpenPool(model, wb);
                LisaOpenPool(model, wb);
                ClosedLoops(model, wb);
                West(model, wb);
                AllOpenPool(model, wb);
                return wb.GetAsByteArray();
            }
        }

        private byte[] EricBiddiscombe(Model model)
        {
            using (ExcelPackage wb = new ExcelPackage())
            {
                wb.Compression = CompressionLevel.BestCompression;

                Eric(model, wb);
                return wb.GetAsByteArray();
            }
        }

        private byte[] LightReport(Model model)
        {
            using (ExcelPackage wb = new ExcelPackage())
            {
                wb.Compression = CompressionLevel.BestCompression;

                AllCustomersLight(model, wb);
                return wb.GetAsByteArray();
            }
        }

        #endregion

        #region Report Parts

        private void EastOpenPool(Model model, ExcelPackage wb)
        {
            var eastOpenPool = model.EastCustomerSummary
                .Join(model.CustomersOpenPool, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .OrderBy(d => d.Row_No_Customer_Data_Type);
            CustomerSummary("East Open Pool", eastOpenPool, wb, Color.FromArgb(84, 130, 53));

            var eastOpenPoolDetails = model.EastDetails
                .Join(model.CustomersOpenPool, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .OrderBy(d => d.Row_No_Customer_Product_Data_Type);
            Deatils("East Open Pool Details", eastOpenPoolDetails, wb, Color.FromArgb(84, 130, 53));
        }

        private void TonyOpenPool(Model model, ExcelPackage wb)
        {
            EastSalesRepOpenPool(model, wb, SalesRepCodes.Tony_Mosco, "Tony");
        }

        /*private void DanaOpenPool(Model model, ExcelPackage wb)
        {
            EastSalesRepOpenPool(model, wb, SalesRepCodes.Dana_Embry, "Dana");
        }

        private void BrandonOpenPool(Model model, ExcelPackage wb)
        {
            EastSalesRepOpenPool(model, wb, SalesRepCodes.Brandon_Osteen, "Brandon");
        }*/

        private void FrankOpenPool(Model model, ExcelPackage wb)
        {
            EastSalesRepOpenPool(model, wb, SalesRepCodes.Frank_Niolet, "Frank");
        }

        private void EastSalesRepOpenPool(Model model, ExcelPackage wb, SalesRepCodes salesRepCode, string salesRepName)
        {
            var openPool = model.EastCustomerSummary
                .Join(model.CustomersOpenPool, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .Where(d => d.Office_Sale_Code == (int)salesRepCode)
                .OrderBy(d => d.Row_No_Customer_Data_Type);
            CustomerSummary(salesRepName + " Open Pool", openPool, wb, Color.FromArgb(169, 208, 142));

            var openPoolDetails = model.EastDetails
                .Join(model.CustomersOpenPool, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .Where(d => d.Office_Sale_Code == (int)salesRepCode)
                .OrderBy(d => d.Row_No_Customer_Product_Data_Type);
            Deatils(salesRepName + " Open Pool Details", openPoolDetails, wb, Color.FromArgb(169, 208, 142));
        }

        /*private void OpenTerritoryOpenPool(Model model, ExcelPackage wb)
        {
            var openTerritoryOpenPool = model.CustomerSummary
                .Join(model.CustomersOpenPool, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .Where(d => d.Office_Sale_Code == (int)SalesRepCodes.Susan_Rivera || model.EastSalesReps.Any(x => x.Office_Sale_Code == d.Office_Sale_Code) == false)
                .Where(d => d.Office_Sale_Code == (int)SalesRepCodes.Susan_Rivera || model.WestSalesReps.Any(x => x.Office_Sale_Code == d.Office_Sale_Code) == false)
                .OrderBy(d => d.Row_No_Customer_Data_Type);
            CustomerSummary("Open Territory Open Pool", openTerritoryOpenPool, wb, Color.FromArgb(84, 130, 53));

            var openTerritoryOpenPoolDetails = model.Details
                .Join(model.CustomersOpenPool, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .Where(d => d.Office_Sale_Code == (int)SalesRepCodes.Susan_Rivera || model.EastSalesReps.Any(x => x.Office_Sale_Code == d.Office_Sale_Code) == false)
                .Where(d => d.Office_Sale_Code == (int)SalesRepCodes.Susan_Rivera || model.WestSalesReps.Any(x => x.Office_Sale_Code == d.Office_Sale_Code) == false)
                .OrderBy(d => d.Row_No_Customer_Product_Data_Type);
            Deatils("Open Territory Open Pool Details", openTerritoryOpenPoolDetails, wb, Color.FromArgb(84, 130, 53));
        }*/

        private void EastLyVsTYBySize(Model model, ExcelPackage wb)
        {
            LyVsTYBySize("East LY vs. TY By Size", model.EastLyVsTYBySize.OrderBy(d => d.Row_No_Product_Data_Type), wb, Color.FromArgb(84, 130, 53));
        }

        private void DedicatedLoops(Model model, ExcelPackage wb)
        {
            var dedicatedLoops = model.CustomerSummary
                .Join(model.CustomersDedicatedLoops, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .OrderBy(d => d.Row_No_Customer_Data_Type);
            CustomerSummary("Dedicated Loops", dedicatedLoops, wb, Color.FromArgb(84, 130, 53));

            var dedicatedLoopsDetails = model.Details
                .Join(model.CustomersDedicatedLoops, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .OrderBy(d => d.Row_No_Customer_Product_Data_Type);
            Deatils("Dedicated Loops Details", dedicatedLoopsDetails, wb, Color.FromArgb(84, 130, 53));
        }

        private void Travis(Model model, ExcelPackage wb)
        {
            var inHouseSales = model.CustomerSummary
                .Join(model.CustomersInsideSales, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .Where(d => d.Office_Sale_Code == (int)SalesRepCodes.Travis_White)
                .OrderBy(d => d.Row_No_Customer_Data_Type);
            CustomerSummary("Travis", inHouseSales, wb, Color.FromArgb(84, 130, 53));

            var inHouseSalesDetails = model.Details
                .Join(model.CustomersInsideSales, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .Where(d => d.Office_Sale_Code == (int)SalesRepCodes.Travis_White)
                .OrderBy(d => d.Row_No_Customer_Product_Data_Type);
            Deatils("Travis Details", inHouseSalesDetails, wb, Color.FromArgb(84, 130, 53));
        }

        private void East(Model model, ExcelPackage wb)
        {
            CustomerSummary("East", model.EastCustomerSummary.OrderBy(d => d.Row_No_Customer_Data_Type), wb, Color.FromArgb(84, 130, 53));
        }

        private void WestOpenPool(Model model, ExcelPackage wb)
        {
            var westOpenPool = model.WestCustomerSummary
                .Join(model.CustomersOpenPool, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .OrderBy(d => d.Row_No_Customer_Data_Type);
            CustomerSummary("West Open Pool", westOpenPool, wb, Color.FromArgb(0, 112, 192));

            var westOpenPoolDetails = model.WestDetails
                .Join(model.CustomersOpenPool, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .OrderBy(d => d.Row_No_Customer_Product_Data_Type);
            Deatils("West Open Pool Details", westOpenPoolDetails, wb, Color.FromArgb(0, 112, 192));
        }

        private void SarahOpenPool(Model model, ExcelPackage wb)
        {
            WestSalesRepOpenPool(model, wb, SalesRepCodes.Sarah_Gonzalez, "Sarah");
        }

        private void GabeOpenPool(Model model, ExcelPackage wb)
        {
            WestSalesRepOpenPool(model, wb, SalesRepCodes.Gabriel_Grijalva, "Gabe");
        }

        private void ConnieOpenPool(Model model, ExcelPackage wb)
        {
            WestSalesRepOpenPool(model, wb, SalesRepCodes.Connie_Ceniceros, "Connie");
        }

        private void LisaOpenPool(Model model, ExcelPackage wb)
        {
            WestSalesRepOpenPool(model, wb, SalesRepCodes.Lisa_Dennis, "Lisa");
        }

        private void WestSalesRepOpenPool(Model model, ExcelPackage wb, SalesRepCodes salesRepCode, string salesRepName)
        {
            var openPool = model.WestCustomerSummary
                .Join(model.CustomersOpenPool, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .Where(d => d.Office_Sale_Code == (int)salesRepCode)
                .OrderBy(d => d.Row_No_Customer_Data_Type);
            CustomerSummary(salesRepName + " Open Pool", openPool, wb, Color.FromArgb(189, 215, 238));

            var openPoolDetails = model.WestDetails
                .Join(model.CustomersOpenPool, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .Where(d => d.Office_Sale_Code == (int)salesRepCode)
                .OrderBy(d => d.Row_No_Customer_Product_Data_Type);
            Deatils(salesRepName + " Open Pool Details", openPoolDetails, wb, Color.FromArgb(189, 215, 238));
        }

        private void WestLyVsTYBySize(Model model, ExcelPackage wb)
        {
            LyVsTYBySize("West LY vs. TY By Size", model.WestLyVsTYBySize.OrderBy(d => d.Row_No_Product_Data_Type), wb, Color.FromArgb(0, 112, 192));
        }

        private void ClosedLoops(Model model, ExcelPackage wb)
        {
            var closedLoops = model.CustomerSummary
                .Join(model.CustomersClosedLoops, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .OrderBy(d => d.Row_No_Customer_Data_Type);
            CustomerSummary("Closed Loops", closedLoops, wb, Color.FromArgb(0, 112, 192));

            var closedLoopsDetails = model.Details
                .Join(model.CustomersClosedLoops, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .OrderBy(d => d.Row_No_Customer_Product_Data_Type);
            Deatils("Closed Loops Details", closedLoopsDetails, wb, Color.FromArgb(0, 112, 192));
        }

        private void West(Model model, ExcelPackage wb)
        {
            CustomerSummary("West", model.WestCustomerSummary.OrderBy(d => d.Row_No_Customer_Data_Type), wb, Color.FromArgb(0, 112, 192));
        }

        private void Meat(Model model, ExcelPackage wb)
        {
            var meat = model.CustomerSummary
                .Join(model.CustomersMeat, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .OrderBy(d => d.Row_No_Customer_Data_Type);
            CustomerSummary("Meat", meat, wb, Color.FromArgb(109, 211, 237));

            var meatDetails = model.Details
                .Join(model.CustomersMeat, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .OrderBy(d => d.Row_No_Customer_Product_Data_Type);
            Deatils("Meat Details", meatDetails, wb, Color.FromArgb(109, 211, 237));
        }

        private void Eggs(Model model, ExcelPackage wb)
        {
            var eggsDetails = model.Details
                .Where(d => d.Product_Code == 1069 || d.Product_Code == 1068)
                .OrderBy(d => d.Row_No_Customer_Product_Data_Type);
            Deatils("Eggs", eggsDetails, wb, Color.FromArgb(182, 105, 118));
        }

        private void Others(Model model, ExcelPackage wb)
        {
            var travisOpenPool = model.EastCustomerSummary
                .Join(model.CustomersOpenPool, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .Where(d => d.Office_Sale_Code == (int)SalesRepCodes.Travis_White)
                .OrderBy(d => d.Row_No_Customer_Data_Type);
            var others = model.CustomerSummary
                .Join(model.CustomersOthers, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .Concat(travisOpenPool)
                .OrderBy(d => d.Row_No_Customer_Data_Type);
            CustomerSummary("Others", others, wb, Color.FromArgb(255, 0, 0));

            var travisOpenPoolDetails = model.EastDetails
                .Join(model.CustomersOpenPool, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .Where(d => d.Office_Sale_Code == (int)SalesRepCodes.Travis_White)
                .OrderBy(d => d.Row_No_Customer_Product_Data_Type);
            var othersDetails = model.Details
                .Join(model.CustomersOthers, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .Concat(travisOpenPoolDetails)
                .OrderBy(d => d.Row_No_Customer_Product_Data_Type);
            Deatils("Others Details", othersDetails, wb, Color.FromArgb(255, 0, 0));
        }

        private void Eric(Model model, ExcelPackage wb)
        {
            SalesRep(model, wb, SalesRepCodes.Eric_Biddiscombe, "Eric");
        }

        private void SalesRep(Model model, ExcelPackage wb, SalesRepCodes salesRepCode, string salesRepName)
        {
            var summary = model.CustomerSummary
                .Where(d => d.Office_Sale_Code == (int)salesRepCode)
                .OrderBy(d => d.Row_No_Customer_Data_Type);
            CustomerSummary(salesRepName, summary, wb, Color.FromArgb(255, 100, 100));

            var details = model.Details
                .Where(d => d.Office_Sale_Code == (int)salesRepCode)
                .OrderBy(d => d.Row_No_Customer_Product_Data_Type);
            Deatils(salesRepName + " Details", details, wb, Color.FromArgb(255, 100, 100));
        }

        private void AllOpenPool(Model model, ExcelPackage wb)
        {
            var openPool = model.CustomerSummary
                .Join(model.CustomersOpenPool, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .OrderBy(d => d.Row_No_Customer_Data_Type);
            CustomerSummary("All Open Pool Customers", openPool, wb, Color.FromArgb(237, 125, 49));

            var openPoolDetails = model.Details
                .Join(model.CustomersOpenPool, d => d.Customer_Code, x => x.Customer_Code, (d, x) => d)
                .OrderBy(d => d.Row_No_Customer_Product_Data_Type);
            Deatils("All Open Pool Customers Details", openPoolDetails, wb, Color.FromArgb(237, 125, 49));
        }

        private void AllLyVsTYBySize(Model model, ExcelPackage wb)
        {
            LyVsTYBySize("All LY vs. TY By Size", model.LyVsTYBySize.OrderBy(d => d.Row_No_Product_Data_Type), wb, Color.FromArgb(234, 102, 58));
        }

        private void AllCustomersLight(Model model, ExcelPackage wb)
        {
            var summary = model.CustomerSummaryLight
                .OrderBy(d => d.Row_No_Customer_Data_Type);
            CustomerSummaryLight("All Customers", summary, wb, Color.FromArgb(237, 125, 49));

            var details = model.Details
                .OrderBy(d => d.Row_No_Customer_Product_Data_Type);
            DeatilsLight("All Customers Details", details, wb, Color.FromArgb(237, 125, 49));
        }

        #endregion

        #region Excel

        private const string textFormat = "@";
        private const string intFormat = "#,##0";
        private const string percentageIntFormat = "0%";
        private const string intFormat0Empty = "#,##0;-#,##0;";
        private const string percentageIntFormat0Empty = "0%;-0%;";

        private const double columnWidthSalesRep = 16;
        private const double columnWidthCustomer = 16;
        private const double columnWidthProduct = 12;
        private const double columnWidthData = 10;
        private const double columnWidthLoop = 12;
        private const double columnWidthCommodity = 12;

        private void CustomerSummary(string tabName, IEnumerable<CustomerSummaryLine> data, ExcelPackage wb, Color tabColor)
        {
            data = data.Concat(GetTotals<CustomerSummaryLine>(data).Select(t => new CustomerSummaryLine()
            {
                Customer_Name = string.Empty,
                Office_Sale_Desc = t.Name,
                Data_Desc = string.Empty,
                M1 = t.M1,
                M2 = t.M2,
                M3 = t.M3,
                M4 = t.M4,
                M5 = t.M5,
                M6 = t.M6,
                M7 = t.M7,
                M8 = t.M8,
                M9 = t.M9,
                M10 = t.M10,
                M11 = t.M11,
                M12 = t.M12,
                Total = t.Total
            }));

            var ws = AddTab<CustomerSummaryLine>(tabName, data, wb, tabColor, GetMembersCustomerSummary());

            int dataCount = data.Count();
            int currentMonth = (dataCount > 0 ? data.First().Current_Month : 0);
            double[] widths = new double[] { columnWidthSalesRep, columnWidthCustomer, columnWidthData };
            FormatTab(ws, headersCustomerSummary, dataCount, currentMonth, widths);
        }

        private void Deatils(string tabName, IEnumerable<DetailsLine> data, ExcelPackage wb, Color tabColor)
        {
            data = data.Concat(GetTotals<DetailsLine>(data).Select(t => new DetailsLine()
            {
                Customer_Name = string.Empty,
                Office_Sale_Desc = t.Name,
                Product_Name = string.Empty,
                Data_Desc = string.Empty,
                M1 = t.M1,
                M2 = t.M2,
                M3 = t.M3,
                M4 = t.M4,
                M5 = t.M5,
                M6 = t.M6,
                M7 = t.M7,
                M8 = t.M8,
                M9 = t.M9,
                M10 = t.M10,
                M11 = t.M11,
                M12 = t.M12,
                Total = t.Total
            }));

            var ws = AddTab<DetailsLine>(tabName, data, wb, tabColor, GetMembersDetails());

            int dataCount = data.Count();
            int currentMonth = (dataCount > 0 ? data.First().Current_Month : 0);
            double[] widths = new double[] { columnWidthSalesRep, columnWidthCustomer, columnWidthProduct, columnWidthData };
            FormatTab(ws, headersDetails, dataCount, currentMonth, widths);
        }

        private void LyVsTYBySize(string tabName, IEnumerable<LYVsTYBySizeLine> data, ExcelPackage wb, Color tabColor)
        {
            data = data.Concat(GetTotals<LYVsTYBySizeLine>(data).Select(t => new LYVsTYBySizeLine()
            {
                Product_Name = t.Name,
                Data_Desc = string.Empty,
                M1 = t.M1,
                M2 = t.M2,
                M3 = t.M3,
                M4 = t.M4,
                M5 = t.M5,
                M6 = t.M6,
                M7 = t.M7,
                M8 = t.M8,
                M9 = t.M9,
                M10 = t.M10,
                M11 = t.M11,
                M12 = t.M12,
                Total = t.Total
            }));

            var ws = AddTab<LYVsTYBySizeLine>(tabName, data, wb, tabColor, GetMembersLyVsTYBySize());

            int dataCount = data.Count();
            int currentMonth = (dataCount > 0 ? data.First().Current_Month : 0);
            double[] widths = new double[] { columnWidthProduct, columnWidthData };
            FormatTab(ws, headersLyVsTYBySize, dataCount, currentMonth, widths);
        }

        private void CustomerSummaryLight(string tabName, IEnumerable<CustomerSummaryLineLight> data, ExcelPackage wb, Color tabColor)
        {
            List<CustomerSummaryLineLight[]> totals = new List<CustomerSummaryLineLight[]>();

            var officeSaleReps = GetOfficeSaleReps(data);
            foreach (var officeSaleRep in officeSaleReps)
            {
                totals.Add(GetTotals<CustomerSummaryLineLight>(data, officeSaleRep).Where(l =>
                    l.Data_Type == (int)DataType.Budget ||
                    l.Data_Type == (int)DataType.Per_Budget ||
                    l.Data_Type == (int)DataType.LY ||
                    l.Data_Type == (int)DataType.TY ||
                    l.Data_Type == (int)DataType.Per_Chg
                ).Select(t => new CustomerSummaryLineLight()
                {
                    Customer_Name = t.Name,
                    Office_Sale_Desc = officeSaleRep,
                    Data_Desc = string.Empty,
                    Loop_Desc = string.Empty,
                    Commodity_Desc = string.Empty,
                    M1 = t.M1,
                    M2 = t.M2,
                    M3 = t.M3,
                    M4 = t.M4,
                    M5 = t.M5,
                    M6 = t.M6,
                    M7 = t.M7,
                    M8 = t.M8,
                    M9 = t.M9,
                    M10 = t.M10,
                    M11 = t.M11,
                    M12 = t.M12,
                    Total = t.Total
                }).ToArray());
            }

            totals.Add(GetTotals<CustomerSummaryLineLight>(data).Where(l =>
                l.Data_Type == (int)DataType.Budget ||
                l.Data_Type == (int)DataType.Per_Budget ||
                l.Data_Type == (int)DataType.LY ||
                l.Data_Type == (int)DataType.TY ||
                l.Data_Type == (int)DataType.Per_Chg
            ).Select(t => new CustomerSummaryLineLight()
            {
                Customer_Name = t.Name,
                Office_Sale_Desc = string.Empty,
                Data_Desc = string.Empty,
                Loop_Desc = string.Empty,
                Commodity_Desc = string.Empty,
                M1 = t.M1,
                M2 = t.M2,
                M3 = t.M3,
                M4 = t.M4,
                M5 = t.M5,
                M6 = t.M6,
                M7 = t.M7,
                M8 = t.M8,
                M9 = t.M9,
                M10 = t.M10,
                M11 = t.M11,
                M12 = t.M12,
                Total = t.Total
            }).ToArray());

            data = data.Concat(totals.SelectMany(t => t));

            data = data.Where(l =>
                l.Data_Type == (int)DataType.Budget ||
                l.Data_Type == (int)DataType.Per_Budget ||
                l.Data_Type == (int)DataType.LY ||
                l.Data_Type == (int)DataType.TY ||
                l.Data_Type == (int)DataType.Per_Chg ||
                l.Data_Desc == string.Empty
            );

            var ws = AddTab<CustomerSummaryLineLight>(tabName, data, wb, tabColor, GetMembersCustomerSummaryLight());

            int dataCount = data.Count();
            int officeSaleRepsCount = officeSaleReps.Count();
            int currentMonth = (dataCount > 0 ? data.First().Current_Month : 0);
            double[] widths = new double[] { columnWidthSalesRep, columnWidthCustomer, columnWidthLoop, columnWidthCommodity, columnWidthData };
            FormatTabLight(ws, headersCustomerSummaryLight, dataCount, officeSaleRepsCount, currentMonth, widths);
        }

        private void DeatilsLight(string tabName, IEnumerable<DetailsLine> data, ExcelPackage wb, Color tabColor)
        {
            var officeSaleReps = GetOfficeSaleReps(data);
            foreach (var officeSaleRep in officeSaleReps)
            {
                data = data.Concat(GetTotals<DetailsLine>(data, officeSaleRep).Where(l =>
                    l.Data_Type == (int)DataType.Budget ||
                    l.Data_Type == (int)DataType.Per_Budget ||
                    l.Data_Type == (int)DataType.LY ||
                    l.Data_Type == (int)DataType.TY ||
                    l.Data_Type == (int)DataType.Per_Chg
                ).Select(t => new DetailsLine()
                {
                    Customer_Name = t.Name,
                    Office_Sale_Desc = officeSaleRep,
                    Product_Name = string.Empty,
                    Data_Desc = string.Empty,
                    Loop_Desc = string.Empty,
                    Commodity_Desc = string.Empty,
                    M1 = t.M1,
                    M2 = t.M2,
                    M3 = t.M3,
                    M4 = t.M4,
                    M5 = t.M5,
                    M6 = t.M6,
                    M7 = t.M7,
                    M8 = t.M8,
                    M9 = t.M9,
                    M10 = t.M10,
                    M11 = t.M11,
                    M12 = t.M12,
                    Total = t.Total
                }));
            }

            data = data.Concat(GetTotals<DetailsLine>(data).Where(l =>
                l.Data_Type == (int)DataType.Budget ||
                l.Data_Type == (int)DataType.Per_Budget ||
                l.Data_Type == (int)DataType.LY ||
                l.Data_Type == (int)DataType.TY ||
                l.Data_Type == (int)DataType.Per_Chg
            ).Select(t => new DetailsLine()
            {
                Customer_Name = t.Name,
                Office_Sale_Desc = string.Empty,
                Product_Name = string.Empty,
                Data_Desc = string.Empty,
                Loop_Desc = string.Empty,
                Commodity_Desc = string.Empty,
                M1 = t.M1,
                M2 = t.M2,
                M3 = t.M3,
                M4 = t.M4,
                M5 = t.M5,
                M6 = t.M6,
                M7 = t.M7,
                M8 = t.M8,
                M9 = t.M9,
                M10 = t.M10,
                M11 = t.M11,
                M12 = t.M12,
                Total = t.Total
            }));

            data = data.Where(l =>
                l.Data_Type == (int)DataType.Budget ||
                l.Data_Type == (int)DataType.Per_Budget ||
                l.Data_Type == (int)DataType.LY ||
                l.Data_Type == (int)DataType.TY ||
                l.Data_Type == (int)DataType.Per_Chg ||
                l.Data_Desc == string.Empty
            );

            var ws = AddTab<DetailsLine>(tabName, data, wb, tabColor, GetMembersDetailsLight());

            int dataCount = data.Count();
            int officeSaleRepsCount = officeSaleReps.Count();
            int currentMonth = (dataCount > 0 ? data.First().Current_Month : 0);
            double[] widths = new double[] { columnWidthSalesRep, columnWidthCustomer, columnWidthLoop, columnWidthCommodity, columnWidthProduct, columnWidthData };
            FormatTabLight(ws, headersDetailsLight, dataCount, officeSaleRepsCount, currentMonth, widths);
        }

        private IEnumerable<string> GetOfficeSaleReps(IEnumerable<IOfficeSaleDesc> data)
        {
            return data
                .Select(osr => osr.Office_Sale_Desc)
                .Distinct()
                .Where(osr =>
                    osr != "Total Forecast" &&
                    osr != "Total % Forecast" &&
                    osr != "Total Budget" &&
                    osr != "Total % Budget" &&
                    osr != "Total LY" &&
                    osr != "Total TY" &&
                    osr != "Total % Chg."
                )
                .OrderBy(osr => osr);
        }

        private void SUMIFS(IEnumerable<string> officeSaleReps, ExcelWorksheet ws)
        {
            int lastRowIndex = ws.Dimension.End.Row;
            int valuesLastRowIndex = lastRowIndex - 7;

            int fromRowIndex = lastRowIndex + 2;
            int toRowIndex = fromRowIndex;
            foreach (var osr in officeSaleReps)
            {
                ws.Cells[toRowIndex, 1].Value = osr;

                for (int sum_range_column = 'D'; sum_range_column <= 'P'; sum_range_column++)
                {
                    int columnIndex = (sum_range_column - 'D') + 4;
                    ws.Cells[toRowIndex, columnIndex].Formula = string.Format(
                        "SUMIFS({0}$2:{0}${1},$A$2:$A${1},$A{2},$C$2:$C${1},\"=TY\")",
                        (char)sum_range_column,
                        valuesLastRowIndex,
                        toRowIndex
                    );
                }

                toRowIndex++;
            }

            toRowIndex--;

            int toColumnIndex = ('P' - 'D') + 4;

            ws.Cells[fromRowIndex, 1, toRowIndex, 1].Style.Numberformat.Format = textFormat;
            ws.Cells[fromRowIndex, 4, toRowIndex, toColumnIndex].Style.Numberformat.Format = intFormat;

            using (var cell = ws.Cells[fromRowIndex, 1, toRowIndex, toColumnIndex])
            {
                cell.Style.Font.Bold = true;
                cell.Style.Font.Color.SetColor(Color.FromArgb(55, 124, 105));
                cell.Style.Fill.PatternType = ExcelFillStyle.None;
            }
        }

        private TotalLine[] GetTotals<T>(IEnumerable<T> data) where T : ILine
        {
            if (data.Count() == 0)
                return new TotalLine[0];

            var forecast = data.Where(d => d.Data_Type == (int)DataType.Forecast);
            var forecastTotal = new TotalLine()
            {
                Name = "Total Forecast",
                Data_Type = (int)DataType.Forecast,
                M1 = forecast.Sum(d => d.M1),
                M2 = forecast.Sum(d => d.M2),
                M3 = forecast.Sum(d => d.M3),
                M4 = forecast.Sum(d => d.M4),
                M5 = forecast.Sum(d => d.M5),
                M6 = forecast.Sum(d => d.M6),
                M7 = forecast.Sum(d => d.M7),
                M8 = forecast.Sum(d => d.M8),
                M9 = forecast.Sum(d => d.M9),
                M10 = forecast.Sum(d => d.M10),
                M11 = forecast.Sum(d => d.M11),
                M12 = forecast.Sum(d => d.M12),
                Total = forecast.Sum(d => d.Total)
            };

            var budget = data.Where(d => d.Data_Type == (int)DataType.Budget);
            var budgetTotal = new TotalLine()
            {
                Name = "Total Budget",
                Data_Type = (int)DataType.Budget,
                M1 = budget.Sum(d => d.M1),
                M2 = budget.Sum(d => d.M2),
                M3 = budget.Sum(d => d.M3),
                M4 = budget.Sum(d => d.M4),
                M5 = budget.Sum(d => d.M5),
                M6 = budget.Sum(d => d.M6),
                M7 = budget.Sum(d => d.M7),
                M8 = budget.Sum(d => d.M8),
                M9 = budget.Sum(d => d.M9),
                M10 = budget.Sum(d => d.M10),
                M11 = budget.Sum(d => d.M11),
                M12 = budget.Sum(d => d.M12),
                Total = budget.Sum(d => d.Total)
            };

            var ly = data.Where(d => d.Data_Type == (int)DataType.LY);
            var lyTotal = new TotalLine()
            {
                Name = "Total LY",
                Data_Type = (int)DataType.LY,
                M1 = ly.Sum(d => d.M1),
                M2 = ly.Sum(d => d.M2),
                M3 = ly.Sum(d => d.M3),
                M4 = ly.Sum(d => d.M4),
                M5 = ly.Sum(d => d.M5),
                M6 = ly.Sum(d => d.M6),
                M7 = ly.Sum(d => d.M7),
                M8 = ly.Sum(d => d.M8),
                M9 = ly.Sum(d => d.M9),
                M10 = ly.Sum(d => d.M10),
                M11 = ly.Sum(d => d.M11),
                M12 = ly.Sum(d => d.M12),
                Total = ly.Sum(d => d.Total)
            };

            var ty = data.Where(d => d.Data_Type == (int)DataType.TY);
            var tyTotal = new TotalLine()
            {
                Name = "Total TY",
                Data_Type = (int)DataType.LY,
                M1 = ty.Sum(d => d.M1),
                M2 = ty.Sum(d => d.M2),
                M3 = ty.Sum(d => d.M3),
                M4 = ty.Sum(d => d.M4),
                M5 = ty.Sum(d => d.M5),
                M6 = ty.Sum(d => d.M6),
                M7 = ty.Sum(d => d.M7),
                M8 = ty.Sum(d => d.M8),
                M9 = ty.Sum(d => d.M9),
                M10 = ty.Sum(d => d.M10),
                M11 = ty.Sum(d => d.M11),
                M12 = ty.Sum(d => d.M12),
                Total = ty.Sum(d => d.Total)
            };

            var perForecastTotal = new TotalLine()
            {
                Name = "Total % Forecast",
                Data_Type = (int)DataType.Per_Forecast,
                M1 = (forecastTotal.M1 != 0.0 ? tyTotal.M1 / forecastTotal.M1 : 1.0),
                M2 = (forecastTotal.M2 != 0.0 ? tyTotal.M2 / forecastTotal.M2 : 1.0),
                M3 = (forecastTotal.M3 != 0.0 ? tyTotal.M3 / forecastTotal.M3 : 1.0),
                M4 = (forecastTotal.M4 != 0.0 ? tyTotal.M4 / forecastTotal.M4 : 1.0),
                M5 = (forecastTotal.M5 != 0.0 ? tyTotal.M5 / forecastTotal.M5 : 1.0),
                M6 = (forecastTotal.M6 != 0.0 ? tyTotal.M6 / forecastTotal.M6 : 1.0),
                M7 = (forecastTotal.M7 != 0.0 ? tyTotal.M7 / forecastTotal.M7 : 1.0),
                M8 = (forecastTotal.M8 != 0.0 ? tyTotal.M8 / forecastTotal.M8 : 1.0),
                M9 = (forecastTotal.M9 != 0.0 ? tyTotal.M9 / forecastTotal.M9 : 1.0),
                M10 = (forecastTotal.M10 != 0.0 ? tyTotal.M10 / forecastTotal.M10 : 1.0),
                M11 = (forecastTotal.M11 != 0.0 ? tyTotal.M11 / forecastTotal.M11 : 1.0),
                M12 = (forecastTotal.M12 != 0.0 ? tyTotal.M12 / forecastTotal.M12 : 1.0),
                Total = (forecastTotal.Total != 0.0 ? tyTotal.Total / forecastTotal.Total : 1.0)
            };

            var perBudgetTotal = new TotalLine()
            {
                Name = "Total % Budget",
                Data_Type = (int)DataType.Per_Budget,
                M1 = (budgetTotal.M1 != 0.0 ? tyTotal.M1 / budgetTotal.M1 : 1.0),
                M2 = (budgetTotal.M2 != 0.0 ? tyTotal.M2 / budgetTotal.M2 : 1.0),
                M3 = (budgetTotal.M3 != 0.0 ? tyTotal.M3 / budgetTotal.M3 : 1.0),
                M4 = (budgetTotal.M4 != 0.0 ? tyTotal.M4 / budgetTotal.M4 : 1.0),
                M5 = (budgetTotal.M5 != 0.0 ? tyTotal.M5 / budgetTotal.M5 : 1.0),
                M6 = (budgetTotal.M6 != 0.0 ? tyTotal.M6 / budgetTotal.M6 : 1.0),
                M7 = (budgetTotal.M7 != 0.0 ? tyTotal.M7 / budgetTotal.M7 : 1.0),
                M8 = (budgetTotal.M8 != 0.0 ? tyTotal.M8 / budgetTotal.M8 : 1.0),
                M9 = (budgetTotal.M9 != 0.0 ? tyTotal.M9 / budgetTotal.M9 : 1.0),
                M10 = (budgetTotal.M10 != 0.0 ? tyTotal.M10 / budgetTotal.M10 : 1.0),
                M11 = (budgetTotal.M11 != 0.0 ? tyTotal.M11 / budgetTotal.M11 : 1.0),
                M12 = (budgetTotal.M12 != 0.0 ? tyTotal.M12 / budgetTotal.M12 : 1.0),
                Total = (budgetTotal.Total != 0.0 ? tyTotal.Total / budgetTotal.Total : 1.0)
            };

            var perChgTotal = new TotalLine()
            {
                Name = "Total % Chg.",
                Data_Type = (int)DataType.Per_Chg,
                M1 = (lyTotal.M1 != 0.0 ? tyTotal.M1 / lyTotal.M1 : 1.0),
                M2 = (lyTotal.M2 != 0.0 ? tyTotal.M2 / lyTotal.M2 : 1.0),
                M3 = (lyTotal.M3 != 0.0 ? tyTotal.M3 / lyTotal.M3 : 1.0),
                M4 = (lyTotal.M4 != 0.0 ? tyTotal.M4 / lyTotal.M4 : 1.0),
                M5 = (lyTotal.M5 != 0.0 ? tyTotal.M5 / lyTotal.M5 : 1.0),
                M6 = (lyTotal.M6 != 0.0 ? tyTotal.M6 / lyTotal.M6 : 1.0),
                M7 = (lyTotal.M7 != 0.0 ? tyTotal.M7 / lyTotal.M7 : 1.0),
                M8 = (lyTotal.M8 != 0.0 ? tyTotal.M8 / lyTotal.M8 : 1.0),
                M9 = (lyTotal.M9 != 0.0 ? tyTotal.M9 / lyTotal.M9 : 1.0),
                M10 = (lyTotal.M10 != 0.0 ? tyTotal.M10 / lyTotal.M10 : 1.0),
                M11 = (lyTotal.M11 != 0.0 ? tyTotal.M11 / lyTotal.M11 : 1.0),
                M12 = (lyTotal.M12 != 0.0 ? tyTotal.M12 / lyTotal.M12 : 1.0),
                Total = (lyTotal.Total != 0.0 ? tyTotal.Total / lyTotal.Total : 1.0)
            };

            return new TotalLine[] {
                forecastTotal,
                perForecastTotal,
                budgetTotal,
                perBudgetTotal,
                lyTotal,
                tyTotal,
                perChgTotal
            };
        }

        private TotalLine[] GetTotals<T>(IEnumerable<T> data, string officeSaleRep) where T : ILineOfficeSaleDesc
        {
            if (data.Count() == 0)
                return new TotalLine[0];

            return GetTotals(
                data.Where(d => d.Office_Sale_Desc == officeSaleRep).Cast<ILine>()
            );
        }

        #region Add Tab

        private ExcelWorksheet AddTab<T>(string tabName, IEnumerable<T> data, ExcelPackage wb, Color tabColor, MemberInfo[] members)
        {
            tabName = tabName.Replace("'", string.Empty);
            var ws = wb.Workbook.Worksheets.Add(tabName);
            ws.View.ShowGridLines = false;
            ws.TabColor = tabColor;
            ws.Cells.Style.Font.Name = "Tahoma";
            ws.Cells.Style.Font.Size = 10;

            using (ExcelRangeBase range = ws.Cells[1, 1].LoadFromCollection<T>(data, true, TableStyles.Light1, memberFlags, members)) { }

            return ws;
        }

        private const BindingFlags memberFlags = BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty;

        private string[] headersCustomerSummary = new string[] { "Sales Rep.", "Customer", "Data", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Total" };

        private MemberInfo[] membersCustomerSummary;
        private MemberInfo[] GetMembersCustomerSummary()
        {
            if (membersCustomerSummary == null)
            {
                List<MemberInfo> members = new List<MemberInfo>();
                MemberInfo[] allMemebers = typeof(CustomerSummaryLine).GetMembers(memberFlags);
                string[] columns = new string[] { "Office_Sale_Desc", "Customer_Name", "Data_Desc", "M1", "M2", "M3", "M4", "M5", "M6", "M7", "M8", "M9", "M10", "M11", "M12", "Total" };
                foreach (var column in columns)
                    members.Add(allMemebers.First(m => m.Name == column));
                membersCustomerSummary = members.ToArray();
            }
            return membersCustomerSummary;
        }

        private string[] headersCustomerSummaryLight = new string[] { "Sales Rep.", "Customer", "Loop", "Commodity", "Data", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Total" };

        private MemberInfo[] membersCustomerSummaryLight;
        private MemberInfo[] GetMembersCustomerSummaryLight()
        {
            if (membersCustomerSummaryLight == null)
            {
                List<MemberInfo> members = new List<MemberInfo>();
                MemberInfo[] allMemebers = typeof(CustomerSummaryLineLight).GetMembers(memberFlags);
                string[] columns = new string[] { "Office_Sale_Desc", "Customer_Name", "Loop_Desc", "Commodity_Desc", "Data_Desc", "M1", "M2", "M3", "M4", "M5", "M6", "M7", "M8", "M9", "M10", "M11", "M12", "Total" };
                foreach (var column in columns)
                    members.Add(allMemebers.First(m => m.Name == column));
                membersCustomerSummaryLight = members.ToArray();
            }
            return membersCustomerSummaryLight;
        }

        private string[] headersDetails = new string[] { "Sales Rep.", "Customer", "Product", "Data", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Total" };

        private MemberInfo[] membersDetails;
        private MemberInfo[] GetMembersDetails()
        {
            if (membersDetails == null)
            {
                List<MemberInfo> members = new List<MemberInfo>();
                MemberInfo[] allMemebers = typeof(DetailsLine).GetMembers(memberFlags);
                string[] columns = new string[] { "Office_Sale_Desc", "Customer_Name", "Product_Name", "Data_Desc", "M1", "M2", "M3", "M4", "M5", "M6", "M7", "M8", "M9", "M10", "M11", "M12", "Total" };
                foreach (var column in columns)
                    members.Add(allMemebers.First(m => m.Name == column));
                membersDetails = members.ToArray();
            }
            return membersDetails;
        }

        private string[] headersDetailsLight = new string[] { "Sales Rep.", "Customer", "Loop", "Commodity", "Product", "Data", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Total" };

        private MemberInfo[] membersDetailsLight;
        private MemberInfo[] GetMembersDetailsLight()
        {
            if (membersDetailsLight == null)
            {
                List<MemberInfo> members = new List<MemberInfo>();
                MemberInfo[] allMemebers = typeof(DetailsLine).GetMembers(memberFlags);
                string[] columns = new string[] { "Office_Sale_Desc", "Customer_Name", "Loop_Desc", "Commodity_Desc", "Product_Name", "Data_Desc", "M1", "M2", "M3", "M4", "M5", "M6", "M7", "M8", "M9", "M10", "M11", "M12", "Total" };
                foreach (var column in columns)
                    members.Add(allMemebers.First(m => m.Name == column));
                membersDetailsLight = members.ToArray();
            }
            return membersDetailsLight;
        }

        private string[] headersLyVsTYBySize = new string[] { "Product", "Data", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Total" };

        private MemberInfo[] membersLyVsTYBySize;
        private MemberInfo[] GetMembersLyVsTYBySize()
        {
            if (membersLyVsTYBySize == null)
            {
                List<MemberInfo> members = new List<MemberInfo>();
                MemberInfo[] allMemebers = typeof(LYVsTYBySizeLine).GetMembers(memberFlags);
                string[] columns = new string[] { "Product_Name", "Data_Desc", "M1", "M2", "M3", "M4", "M5", "M6", "M7", "M8", "M9", "M10", "M11", "M12", "Total" };
                foreach (var column in columns)
                    members.Add(allMemebers.First(m => m.Name == column));
                membersLyVsTYBySize = members.ToArray();
            }
            return membersLyVsTYBySize;
        }

        #endregion

        #region Formatting

        private void FormatTab(ExcelWorksheet ws, string[] headers, int dataCount, int currentMonth, double[] widths)
        {
            int textColumnCount = widths.Length;

            FormatHeaders(ws, headers);
            FormatCells(ws, dataCount, textColumnCount, currentMonth);
            SetColumnWidths(ws, widths);
            ws.View.FreezePanes(2, 1);

            ExcelTable tblData = ws.Tables[ws.Tables.Count - 1];
            tblData.Name = "tbl" + ws.Name.Replace(" ", string.Empty).Replace(".", string.Empty);

            tblData.EnableFilters(
                filterColumnIndex: textColumnCount,
                filterValues: new string[] { "LY", "TY", "% Chg." },
                blanks: true
            );
        }

        private void FormatTabLight(ExcelWorksheet ws, string[] headers, int dataCount, int officeSaleRepsCount, int currentMonth, double[] widths)
        {
            int textColumnCount = widths.Length;

            FormatHeaders(ws, headers);
            FormatCellsLight(ws, dataCount, officeSaleRepsCount, textColumnCount, currentMonth);
            SetColumnWidths(ws, widths);
            ws.View.FreezePanes(2, 1);

            ExcelTable tblData = ws.Tables[ws.Tables.Count - 1];
            tblData.Name = "tbl" + ws.Name.Replace(" ", string.Empty).Replace(".", string.Empty);
        }

        private void FormatHeaders(ExcelWorksheet ws, string[] headers)
        {
            for (int i = 0; i < headers.Length; i++)
                ws.Cells[1, i + 1].Value = headers[i];

            using (var cells = ws.Cells[1, 1, 1, headers.Length])
            {
                cells.Style.Font.Size = 11;
                cells.Style.Font.Bold = true;
                cells.Style.Font.Color.SetColor(Color.White);
                cells.Style.Numberformat.Format = textFormat;
                cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cells.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.LightGray);
                cells.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cells.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(112, 173, 71));
            }
        }

        private void FormatCells(ExcelWorksheet ws, int dataCount, int textColumnCount, int currentMonth)
        {
            if (dataCount == 0)
                return;

            ws.Cells[2, 1, dataCount + 1, textColumnCount].Style.Numberformat.Format = textFormat;
            ws.Cells[2, textColumnCount + 1, dataCount + 1, textColumnCount + 1 + 12].Style.Numberformat.Format = intFormat0Empty;

            SetPercentageCellsFormat(ws, dataCount, textColumnCount, 3);
            SetPercentageCellsFormat(ws, dataCount, textColumnCount, 5);
            SetPercentageCellsFormat(ws, dataCount, textColumnCount, 8);

            int fromMonthColumnIndex = textColumnCount + 1;
            int toMonthColumnIndex = textColumnCount + 12;
            int currentMonthColumnIndex = fromMonthColumnIndex + (currentMonth - 1);
            int nextMonthColumnIndex = currentMonthColumnIndex + 1;
            if (nextMonthColumnIndex <= toMonthColumnIndex)
            {
                for (int rowIndex = 7; rowIndex <= dataCount + 1; rowIndex += 7)
                    ws.Cells[rowIndex, nextMonthColumnIndex, rowIndex, toMonthColumnIndex].Style.Font.Color.SetColor(Color.LimeGreen);
            }

            int fromTotalsRowIndex = dataCount - 5;
            int toTotalsRowIndex = fromTotalsRowIndex + 6;

            using (var cells = ws.Cells[fromTotalsRowIndex, 1, toTotalsRowIndex, textColumnCount + 1 + 12])
            {
                cells.Style.Font.Bold = true;
                cells.Style.Font.Color.SetColor(Color.FromArgb(55, 86, 35));
                cells.Style.Fill.PatternType = ExcelFillStyle.None;
            }

            ws.Cells[fromTotalsRowIndex, 1, toTotalsRowIndex, textColumnCount].Style.Numberformat.Format = textFormat;
            ws.Cells[fromTotalsRowIndex, textColumnCount + 1, toTotalsRowIndex, textColumnCount + 1 + 12].Style.Numberformat.Format = intFormat;
            ws.Cells[fromTotalsRowIndex + 1, textColumnCount + 1, fromTotalsRowIndex + 1, textColumnCount + 1 + 12].Style.Numberformat.Format = percentageIntFormat;
            ws.Cells[fromTotalsRowIndex + 3, textColumnCount + 1, fromTotalsRowIndex + 3, textColumnCount + 1 + 12].Style.Numberformat.Format = percentageIntFormat;
            ws.Cells[fromTotalsRowIndex + 6, textColumnCount + 1, fromTotalsRowIndex + 6, textColumnCount + 1 + 12].Style.Numberformat.Format = percentageIntFormat;
        }

        private void FormatCellsLight(ExcelWorksheet ws, int dataCount, int officeSaleRepsCount, int textColumnCount, int currentMonth)
        {
            if (dataCount == 0)
                return;

            int detailsStep = 5;
            int subtotalsStep = 5;
            int totalsStep = 5;

            int totalsToRowIndex = dataCount + 1;
            int totalsFromRowIndex = totalsToRowIndex - (totalsStep - 1);
            int subtotalsToRowIndex = totalsFromRowIndex - 1;
            int subtotalsFromRowIndex = subtotalsToRowIndex - (subtotalsStep * officeSaleRepsCount) + 1;
            int detailsToRowIndex = subtotalsFromRowIndex - 1;
            int detailsFromRowIndex = 2;

            int detailsCount = detailsToRowIndex - detailsFromRowIndex + 1;
            int subtotalsCount = subtotalsToRowIndex - subtotalsFromRowIndex + 1;

            ws.Cells[detailsFromRowIndex, 1, totalsToRowIndex, textColumnCount].Style.Numberformat.Format = textFormat;
            ws.Cells[detailsFromRowIndex, textColumnCount + 1, detailsToRowIndex, textColumnCount + 1 + 12].Style.Numberformat.Format = intFormat0Empty;
            ws.Cells[subtotalsFromRowIndex, textColumnCount + 1, totalsToRowIndex, textColumnCount + 1 + 12].Style.Numberformat.Format = intFormat;

            SetPercentageCellsFormat(ws, detailsToRowIndex - 1, textColumnCount, 3, detailsStep); // % Budget
            SetPercentageCellsFormat(ws, detailsToRowIndex - 1, textColumnCount, 6, detailsStep); // % Chg.
            SetPercentageCellsFormat(ws, subtotalsToRowIndex - 1, textColumnCount, subtotalsFromRowIndex + 1, subtotalsStep, percentageIntFormat); // Total % Budget
            SetPercentageCellsFormat(ws, subtotalsToRowIndex - 1, textColumnCount, subtotalsFromRowIndex + 4, subtotalsStep, percentageIntFormat); // Total % Chg.
            SetPercentageCellsFormat(ws, totalsToRowIndex - 1, textColumnCount, totalsFromRowIndex + 1, totalsStep, percentageIntFormat); // Total % Budget
            SetPercentageCellsFormat(ws, totalsToRowIndex - 1, textColumnCount, totalsFromRowIndex + 4, totalsStep, percentageIntFormat); // Total % Chg.

            int fromMonthColumnIndex = textColumnCount + 1;
            int toMonthColumnIndex = textColumnCount + 12;
            int currentMonthColumnIndex = fromMonthColumnIndex + (currentMonth - 1);
            int nextMonthColumnIndex = currentMonthColumnIndex + 1;
            if (nextMonthColumnIndex <= toMonthColumnIndex)
            {
                for (int rowIndex = 5; rowIndex <= detailsToRowIndex; rowIndex += detailsStep)
                    ws.Cells[rowIndex, nextMonthColumnIndex, rowIndex, toMonthColumnIndex].Style.Font.Color.SetColor(Color.LimeGreen);
            }

            using (var cells = ws.Cells[subtotalsFromRowIndex, 1, subtotalsToRowIndex, textColumnCount + 1 + 12])
            {
                cells.Style.Font.Bold = true;
                cells.Style.Font.Color.SetColor(Color.FromArgb(55, 124, 105));
                cells.Style.Fill.PatternType = ExcelFillStyle.None;
            }

            using (var cells = ws.Cells[totalsFromRowIndex, 1, totalsToRowIndex, textColumnCount + 1 + 12])
            {
                cells.Style.Font.Bold = true;
                cells.Style.Font.Color.SetColor(Color.FromArgb(55, 86, 35));
                cells.Style.Fill.PatternType = ExcelFillStyle.None;
            }
        }

        private void SetPercentageCellsFormat(ExcelWorksheet ws, int dataCount, int textColumnCount, int rowIndex, int step = 7, string percentageFormat = percentageIntFormat0Empty)
        {
            for (; rowIndex <= dataCount + 1; rowIndex += step)
            {
                ws.Cells[rowIndex, 1, rowIndex, textColumnCount + 1 + 12].Style.Font.Bold = true;
                ws.Cells[rowIndex, textColumnCount + 1, rowIndex, textColumnCount + 1 + 12].Style.Numberformat.Format = percentageFormat;
            }
        }

        private void SetColumnWidths(ExcelWorksheet ws, params double[] widths)
        {
            int columnIndex = 1;
            for (; columnIndex - 1 < widths.Length; columnIndex++)
                ws.Column(columnIndex).Width = widths[columnIndex - 1];

            for (; columnIndex <= ws.Dimension.End.Column; columnIndex++)
            {
                var column = ws.Column(columnIndex);
                column.AutoFit();
                column.Width += 2;
            }
        }

        #endregion

        #endregion
    }
}
