using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using Dapper;

namespace AutoInvoicesUS
{
    class AutoInvoices
    {
        public static readonly bool IsProduction = true;

        public void Run()
        {
            DateTime now = DateTime.Now;

            try
            {
                now = GetSQLServerDateTimeNow();
                RunAutoInvoices(now);
            }
            catch (Exception ex)
            {
                Console.WriteLine("{0:yyyy-MM-dd HH:mm:ss} Unexpected error: {1}", now, ex.Message);
                Console.WriteLine("{0:yyyy-MM-dd HH:mm:ss} Error stack trace: {1}", now, ex.StackTrace);
            }
        }

        private void RunAutoInvoices(DateTime now)
        {
            bool isFailedGetRSSDieselFuel = false;
            Exception rssDieselFuelError = null;
            List<Tuple<string, double>> dieselLines = GetRSSDiesel(ref isFailedGetRSSDieselFuel, ref rssDieselFuelError);
            if (isFailedGetRSSDieselFuel)
            {
                Console.WriteLine("{0:yyyy-MM-dd HH:mm:ss} {1}", now, InvoicingDisabledMessage);
                if (rssDieselFuelError != null && rssDieselFuelError.InnerException != null)
                    Console.WriteLine("{0:yyyy-MM-dd HH:mm:ss} AutoInvoicesUS.exe RSS feed Error: {1}", now, rssDieselFuelError.InnerException.Message);

                if (rssDieselFuelError != null)
                    ExceptionsLogger.AddExceptionToDatabase(rssDieselFuelError, now);

                return;
            }

            var dieselLinesDT = dieselLines.Select(x => new DieselLine() { Region = x.Item1, Diesel = (decimal)x.Item2 }).ToDataTable();
            Tuple<IEnumerable<InvoiceLine>, IEnumerable<FSCMileageLine>, IEnumerable<ResultLine>, IEnumerable<FileStorageLine>> lines = null;
            string autoInvoicesErrorMessage = null;
            bool succeeded = sp_us_accounting_add_to_priority_auto(dieselLinesDT, ref lines, ref autoInvoicesErrorMessage);

            if (string.IsNullOrEmpty(autoInvoicesErrorMessage) == false)
            {
                Console.WriteLine("{0:yyyy-MM-dd HH:mm:ss} Failed to execute auto invoices: {1}", now, autoInvoicesErrorMessage);
                return;
            }

            if (succeeded == false || lines == null || lines.Item3 == null)
            {
                Console.WriteLine("{0:yyyy-MM-dd HH:mm:ss} Auto invoices failed", now);
                return;
            }

            var invoiceLines = lines.Item1;
            var mileageLines = lines.Item2;
            var resultLines = lines.Item3;
            var fileStorageLines = lines.Item4;

            if (resultLines.Count() == 0)
            {
                Console.WriteLine("{0:yyyy-MM-dd HH:mm:ss} No movements were found", now);
                return;
            }

            invoiceLines.Where(x => x.Is_Manual_Calculation == 1).ForEach(x => x.Comments == "Manual calculation - please add tax deduction rate");

            PrintResults(resultLines);

            string invoiceFolder = null;
            string archiveFolder = null;
            string exportPathsErrorMessage = null;
            GetExportFolders(ref invoiceFolder, ref archiveFolder, ref exportPathsErrorMessage);
            if (string.IsNullOrEmpty(exportPathsErrorMessage) == false)
            {
                Console.WriteLine("{0:yyyy-MM-dd HH:mm:ss} Failed to create file export paths: {1}", now, exportPathsErrorMessage);
                return;
            }

            string excelExportErrorMessage = null;
            SaveExcelToFile(invoiceLines, mileageLines, resultLines, now, invoiceFolder, archiveFolder, ref excelExportErrorMessage);
            if (string.IsNullOrEmpty(excelExportErrorMessage) == false)
                Console.WriteLine("{0:yyyy-MM-dd HH:mm:ss} Failed to export excel: {1}", now, excelExportErrorMessage);

            string pdfExportErrorMessage = null;
            SavePDFToFile(invoiceLines, mileageLines, resultLines, now, invoiceFolder, archiveFolder, ref pdfExportErrorMessage);
            if (string.IsNullOrEmpty(pdfExportErrorMessage) == false)
                Console.WriteLine("{0:yyyy-MM-dd HH:mm:ss} Failed to export pdf: {1}", now, pdfExportErrorMessage);

            string noteExportErrorMessage = null;
            SaveScannedNotes(resultLines, fileStorageLines, now, invoiceFolder, archiveFolder, ref noteExportErrorMessage);
            if (string.IsNullOrEmpty(noteExportErrorMessage) == false)
                Console.WriteLine("{0:yyyy-MM-dd HH:mm:ss} Failed to export scanned notes: {1}", now, noteExportErrorMessage);
        }

        private static DateTime GetSQLServerDateTimeNow()
        {
            return SqlHelper.ExecuteScalarSP<DateTime>("sp_get_db_date");
        }

        private string InvoicingDisabledMessage = "AutoInvoicesUS.exe: Failed to retrieve On-Highway Diesel Fuel Retail Price from RSS feed.";

        private List<Tuple<string, double>> GetRSSDiesel(ref bool isFailedGetRSSDieselFuel, ref Exception rssDieselFuelError)
        {
            List<Tuple<string, double>> dieselLines = null;
            Exception rssDieselFuelInnerError = null;

            try
            {
                dieselLines = FuelSurcharge.GetRSSDieselFuel();
            }
            catch (Exception ex)
            {
                dieselLines = null;
                rssDieselFuelInnerError = ex;
            }

            if (dieselLines == null || dieselLines.Count == 0)
            {
                isFailedGetRSSDieselFuel = true;

                if (rssDieselFuelInnerError != null)
                    rssDieselFuelError = new Exception(InvoicingDisabledMessage, rssDieselFuelInnerError);
                else
                    rssDieselFuelError = new Exception(InvoicingDisabledMessage);
            }

            return dieselLines;
        }

        private bool sp_us_accounting_add_to_priority_auto(DataTable dieselLinesDT, ref Tuple<IEnumerable<InvoiceLine>, IEnumerable<FSCMileageLine>, IEnumerable<ResultLine>, IEnumerable<FileStorageLine>> lines, ref string errorMessage)
        {
            try
            {
                var outParam = new DynamicParameters("RETURN_VALUE", sqlDbType: SqlDbType.Int, direction: ParameterDirection.ReturnValue);
                lines = SqlHelper.QueryMultipleSP<InvoiceLine, FSCMileageLine, ResultLine, FileStorageLine>("sp_us_accounting_add_to_priority_auto", new { DIESELS = dieselLinesDT }, outParam);
                int? returnValue = outParam.Get<int?>("RETURN_VALUE");
                bool succeeded = (returnValue == 0);
                return succeeded;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
            }

            lines = null;
            return false;
        }

        private void PrintResults(IEnumerable<ResultLine> resultLines)
        {
            DataTable resultsDT = resultLines.Select(line => new
            {
                line.TTP_Numerator,
                Customer = line.Customer_Name,
                line.Priority_Lines_Count,
                Succeeded = (line.Succeeded == 1 ? "Yes" : "No"),
                line.Priority_Numerator,
                line.Message
            }).ToDataTable();

            foreach (DataColumn column in resultsDT.Columns)
                column.ColumnName = column.ColumnName.Replace('_', ' ');

            PrintDataExtensions.ASCIIBorder();
            resultsDT.PrintList(repeatColumns: 1);
        }

        private void GetExportFolders(ref string invoiceFolder, ref string archiveFolder, ref string errorMessage)
        {
            try
            {
                invoiceFolder = ConfigurationManager.AppSettings["Invoice_Folder"];
                invoiceFolder = invoiceFolder.Trim();
                if (invoiceFolder.EndsWith(Path.DirectorySeparatorChar.ToString()) == false)
                    invoiceFolder += Path.DirectorySeparatorChar;
                if (Directory.Exists(invoiceFolder) == false)
                    Directory.CreateDirectory(invoiceFolder);
                archiveFolder = invoiceFolder + "archive" + Path.DirectorySeparatorChar;
                if (Directory.Exists(archiveFolder) == false)
                    Directory.CreateDirectory(archiveFolder);
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
            }
        }

        private void SaveExcelToFile(IEnumerable<InvoiceLine> invoiceLines, IEnumerable<FSCMileageLine> mileageLines, IEnumerable<ResultLine> resultLines, DateTime now, string invoiceFolder, string archiveFolder, ref string errorMessage)
        {
            try
            {
                foreach (var line in resultLines.Where(line => line.Succeeded == 1))
                {
                    byte[] excel = ExportHelper.ExcelExportHelper.GetExcel(
                        line.TTP_Numerator,
                        line.Customer_Name,
                        invoiceLines.Where(x => x.Numerator == line.TTP_Numerator),
                        mileageLines.Where(x => x.Numerator == line.TTP_Numerator),
                        now
                    );

                    string filePath = invoiceFolder + line.Priority_Numerator + ".xlsx";
                    File.WriteAllBytes(filePath, excel);

                    filePath = archiveFolder + line.Priority_Numerator + ".xlsx";
                    File.WriteAllBytes(filePath, excel);
                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
            }
        }

        private void SavePDFToFile(IEnumerable<InvoiceLine> invoiceLines, IEnumerable<FSCMileageLine> mileageLines, IEnumerable<ResultLine> resultLines, DateTime now, string invoiceFolder, string archiveFolder, ref string errorMessage)
        {
            try
            {
                foreach (var line in resultLines.Where(line => line.Succeeded == 1))
                {
                    byte[] pdf = ExportHelper.PDFExportHelper.GetPDF(
                        line.TTP_Numerator,
                        line.Customer_Name,
                        invoiceLines.Where(x => x.Numerator == line.TTP_Numerator),
                        mileageLines.Where(x => x.Numerator == line.TTP_Numerator),
                        now
                    );

                    string filePath = invoiceFolder + line.Priority_Numerator + ".pdf";
                    File.WriteAllBytes(filePath, pdf);

                    filePath = archiveFolder + line.Priority_Numerator + ".pdf";
                    File.WriteAllBytes(filePath, pdf);
                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
            }
        }

        private void SaveScannedNotes(IEnumerable<ResultLine> resultLines, IEnumerable<FileStorageLine> fileStorageLines, DateTime now, string invoiceFolder, string archiveFolder, ref string errorMessage)
        {
            try
            {
                foreach (var line in resultLines.Where(line => line.Succeeded == 1))
                {
                    foreach (var note in fileStorageLines.Where(x => x.TTP_Numerator == line.TTP_Numerator))
                    {
                        int maxFileNameLength = 40;
                        string File_Name_Desc = note.File_Name_Desc;
                        if (string.IsNullOrEmpty(File_Name_Desc) == false && File_Name_Desc.Length > maxFileNameLength)
                        {
                            int lastDotIndex = File_Name_Desc.LastIndexOf('.');
                            if (lastDotIndex != -1)
                                File_Name_Desc = File_Name_Desc.Substring(0, maxFileNameLength - 4) + File_Name_Desc.Substring(lastDotIndex);
                            else
                                File_Name_Desc = File_Name_Desc.Substring(0, maxFileNameLength);
                        }

                        File_Name_Desc = ExportHelper.FileExportHelper.CleanFileName(File_Name_Desc);

                        string noteFilePath = invoiceFolder + File_Name_Desc;
                        File.WriteAllBytes(noteFilePath, note.File_Stream);
                    }
                }
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
            }
        }
    }
}
