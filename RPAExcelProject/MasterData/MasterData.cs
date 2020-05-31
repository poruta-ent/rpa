using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;

namespace RPAExcelProject
{
    public static class MasterData
    {

        public static string reportDate;
        public static ExcelPackage reportWorkbook;
        public static ExcelWorksheet reportSheet;

        // KK 2019-08-02 data poprzedniego raportu wykorzystywana w SEB
        public static DateTime previousReportDate;

        public static List<string> bankAccounts;
        public static List<string> accountCurrencies;
        
        public static Dictionary<string, CompanyBankAccount> cfReportLines;

        public static void InitializeMasterdata()
        {
            CompanyNamesMapping.GenerateMappings();
            bankAccounts = InitializeBankAccounts();
            accountCurrencies = InitializeCurrencies();
            cfReportLines = InitializeCfReportLines(reportWorkbook);
        }

        
        /// <summary>
        /// Zwraca listę rachunków - bank + typ.
        /// </summary>
        /// <returns></returns>
        //TODO BIZNES potwierdzić listę rachunków, czy możemy trzymać ją w pliku konfiguracyjnym, może przejść na numery rachuków?
        public static List<string> InitializeBankAccounts()
        {
            return new List<string>()
                {
                    "ING CP",
                    "SPLIT",
                    "Escrow ING",
                    "BZ WBK",
                    "SEB SE",
                    "SEB CP",
                    "SEB N",
                    "ING N",
                    "ING Płace",
                    "ING ZFŚS",
                    "Santander",
                    "ING (ZFŚS)",
                    "ING BV",
                    "ING FUND"
                };

        }

        /// <summary>
        /// Zwraca listę walut - so ustalenia jak wygląda kompletna
        /// </summary>
        /// <returns></returns>
        //TODO BIZNES lista używanych walut, może pobierać z K2?
        public static List<string> InitializeCurrencies()
        {
            return new List<string>()
                {
                    "EUR",
                    "PLN",
                };
        }

        /// <summary>
        /// Zwraca kod linii raportowej CF
        /// </summary>
        /// <returns></returns>
        public static string GetCfLineKey(string accountBank, string accountCurrency, string bankCompanyName)
        {
            return $"{accountBank}###{accountCurrency}###{bankCompanyName.ToUpper()}";
        }

        /// <summary>
        /// Zwraca słownik ze szczegółami poszczegolnych linii raportowych - CfReportLines.
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, CompanyBankAccount> InitializeCfReportLines(ExcelPackage reportWorkbook)
        {
            var cfReportLines = new Dictionary<string, CompanyBankAccount>();
            var reportTemplateSheet = reportWorkbook.Workbook.Worksheets["T"];

            int colKey = 4;
            int colAccountDescription = 2;
            int colAccountNumber = 3;

            for (int row = 15; row <= reportTemplateSheet.Dimension.End.Row; row++)
            {
                if (reportTemplateSheet.Cells[row, 2].Value != null)
                {
                    string bankCompanyName = reportTemplateSheet.Cells[row, colKey].Value.ToString().ToUpper();
                    string accountCurrency = reportTemplateSheet.Cells[row, colAccountDescription].Value.ToString().Split(" ")[0];
                    string accountBank = bankAccounts.Where(bank => reportTemplateSheet.Cells[row, colAccountDescription].Value.ToString().Contains(bank)).FirstOrDefault();

                    //TODO KONCEPCJA - czy taki klucz będzie wystarczający? Może przejść na numer rachunku
                    string accountKey = $"{accountBank}###{accountCurrency}###{bankCompanyName}";

                    if (!cfReportLines.ContainsKey(accountKey))
                    {
                        cfReportLines.Add(accountKey,
                                        new CompanyBankAccount()
                                        {
                                            CompanyNameInBank = bankCompanyName,
                                            AccountDesc = reportTemplateSheet.Cells[row, colAccountDescription].Value.ToString(),
                                            RowInCfReport = row,
                                            CompanyNaneInKorab2 = CompanyNamesMapping.CFAsKey.ContainsKey(accountKey) ? CompanyNamesMapping.CFAsKey[accountKey].Korab2Name : string.Empty,
                                            Currency = accountCurrency,
                                            Bank = accountBank,
                                            AccountNumber = reportTemplateSheet.Cells[row, colAccountNumber].Value?.ToString()
                                        }
                                        );
                    }
                }

            }
            return cfReportLines;
        }

        /// <summary>
        /// Tworzy arkusz raportu CF na podstawie templatu.
        /// </summary>
        /// <returns></returns>
        public static ExcelWorksheet CreateFromTemplate()
        {
            string mostRecentReport = MostRecentReportName(reportWorkbook);
            var newCfWorksheet = reportWorkbook.Workbook.Worksheets.Add(reportDate, reportWorkbook.Workbook.Worksheets["T"]);
            reportWorkbook.Workbook.Worksheets.MoveAfter(newCfWorksheet.Name, mostRecentReport);
            //TODO V3 jak sprawić żeby nowy arkusz się aktywował bo poniższe nie działa
            reportWorkbook.Workbook.Worksheets[reportDate].Select();
            reportWorkbook.Save();
            return newCfWorksheet;
        }

        /// <summary>
        /// Zraca nazwę arusza która odpowiada najpóźniejszej dacie.
        /// </summary>
        /// <returns></returns>
        public static string MostRecentReportName(ExcelPackage reportWorkbook)
        {

            List<DateTime> reportDates = new List<DateTime>();
            DateTime reportDate;

            for (int i = 0; i < reportWorkbook.Workbook.Worksheets.Count; i++)
            {
                if (DateTime.TryParse(reportWorkbook.Workbook.Worksheets[i].Name, out reportDate))
                {
                    reportDates.Add(reportDate);
                }
            }

            var lastReportDate = reportDates.OrderByDescending(i => i).FirstOrDefault();

            return lastReportDate.ToString("dd.MM.yyyy");
        }
    }


}
