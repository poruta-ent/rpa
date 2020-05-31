using OfficeOpenXml;
using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;

namespace RPAExcelProject
{

   public static class ReportFile
    {
        #region General
        public static string CFReport => "CF dzienny SPP 09 08 2019.xlsx";
        public static string CompanyMapping => "mapowanie_spó³ki.xlsx";
        public static string OutflowsInflows => "SPP ING 09.08.xlsx";
        #endregion

        #region Bank INGPL
        public static string OpeningBalanceINGPL => "Daily Flows Poland_ROBOT 2_A_20190809065838.xlsx";
        public static string OpeningBalanceTFI => "TFI SALDA_M_20190809073615.xlsx";
        public static string OpeningBalanceFizan => "FIZAN_salda_M_20190809073639.xlsx";
        public static string OutflowsInflowsINGSPP => "SPP ING 09.08.xlsx";
        public static string OutflowsInflowsINGCDE => "CDE ING 09.08.xlsx";
        public static string OutflowsInflowsINGTFI => "TFI ING 09.08.xlsx";
        public static string SplitPaymentINGPL => "SALDA SPLIT_M_20190809073546.xlsx";
        public static string DailyLimitINGPLEUR => "LIMITY DZIENNE_Limity EUR_A_20190809081927.xlsx";
        public static string DailyLimitINGPLPLN => "LIMITY DZIENNE_LIMITY DZIENNE PLN_A_20190809081910.xlsx";
        #endregion

        #region Bank SEB, Santande, ING NL
        // KK Bank SEB wg propozycji Ewy Jeziorek ("SEB_ROBOT_NEW yyyyMMdd.csv")
        public static string OpeningBalanceSEB;
        public static string OpeningBalanceSantander;
        public static string OpeningBalanceINGNL;
        #endregion


        #region IndividualSettings
        public static string FolderPath => @"C:\RPA_FILES\";
        #endregion

        #region Log Sheets
        public static ExcelPackage logWorkbook;
       
        public static bool SplitTest = true;
        public static ExcelWorksheet splitLogs;
        
        #endregion

        /// <summary>
        /// Generuje obiekt ExcelPackage dla pliku znajduj¹cego sie w folderze (lub tworzy nowy plik je¿eli go nie ma).
        /// </summary>
        /// <returns></returns>
        public static ExcelPackage GetWorkbook(string name) //string folder, string file, bool canCreate)
        {
            return new ExcelPackage(new FileInfo(FolderPath + name));
        }

        /// <summary>
        /// Wczytuje plik .CSV i tworzy w pamiêci skoroszyt Excell
        /// </summary> 
        /// <param name="fileName"></param>
        /// <param name="reportDate"></param>
        /// <returns></returns>
        public static ExcelWorksheet GetWorkbookCsv(string fileName, string reportDate, string bankName)
        {
            ExcelTextFormat format = new ExcelTextFormat();
            if (bankName == "SEB" || bankName == "Santander")
                format.Delimiter = ';';
            else
                format.Delimiter = ',';

            format.Culture = new CultureInfo(Thread.CurrentThread.CurrentCulture.ToString());
            format.Culture.DateTimeFormat.ShortDatePattern = "dd-MM-yyyy";
            //KK ile kolumn mo¿e byæ w arkuszu (przyjmuje dowolnie du¿¹)
            format.DataTypes = new eDataTypes[100];

            if (bankName == "Santander")
            {
                format.Encoding = Encoding.GetEncoding(1250);
                format.EOL = "\n";
                //KK okreœlenie, ¿e druga kolumna ma byæ jako tekst, dziêki temu pole zawieraj¹ce nr rachunku w NRB
                //  nie zamienia siê na liczbê
                format.DataTypes[1] = eDataTypes.String;    //druga kolumna (numeracja od 0)
            }
            else if (bankName == "ING BV")
            {
                format.Encoding = new UTF8Encoding();
                format.EOL = "\n";
            }
            else
            {
                //KK SEB ma znaki koñca linii jak w Windows 0D0A
                format.Encoding = new UTF8Encoding();
            }

            FileInfo file = new FileInfo(ReportFile.FolderPath + fileName);
            ExcelPackage excelPackage = new ExcelPackage();

            ExcelWorksheet _worksheet = excelPackage.Workbook.Worksheets.Add($"{reportDate}");
            //load the CSV data into cell A1
            _worksheet.Cells["A1"].LoadFromText(file, format);

            string filePathName = ReportFile.FolderPath + "ZZ_SEB.xlsx";

            if (bankName == "SEB")
            {

                // KK ustawia kolumnê E w formacie daty
                _worksheet.Column(5).Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
            }
            else if (bankName == "Santander")
            {
                filePathName = ReportFile.FolderPath + "ZZ_Santander.xlsx";
            }
            else if (bankName == "ING BV")
            {
                // KK ustawia kolumnê E w formacie daty
                _worksheet.Column(2).Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                filePathName = ReportFile.FolderPath + "ZZ_INGBV.xlsx";
            }

            FileInfo fi = new FileInfo(filePathName);
            excelPackage.SaveAs(fi);

            return _worksheet;
        }

        /// <summary>
        /// Sprawdza, czy w katalogu FolderPath znajduj¹ siê wszystkie potrzebne pliki i czy czasami nie s¹ otwarte.
        /// </summary>
        /// <returns></returns>
        public static bool CheckIfCFFilesReady(out string message)
        {
            StringBuilder missedFiles = new StringBuilder();
            StringBuilder openedFiles = new StringBuilder();
            StringBuilder errMessage = new StringBuilder();

            if (File.Exists(FolderPath + CFReport))
            {
                if (Utils.IsOpen(FolderPath + CFReport)) openedFiles.Append($"{CFReport}\n");
            }
            else
                missedFiles.Append($"{CFReport}\n");

            if (File.Exists(FolderPath + CompanyMapping))
            {
                if (Utils.IsOpen(FolderPath + CompanyMapping)) openedFiles.Append($"{CompanyMapping}\n");
            }
            else
                missedFiles.Append($"{CompanyMapping}\n");

            if (File.Exists(FolderPath + OutflowsInflows))
            {
                if (Utils.IsOpen(FolderPath + OutflowsInflows)) openedFiles.Append($"{OutflowsInflows}\n");
            }
            else
                missedFiles.Append($"{OutflowsInflows}\n");

            if (File.Exists(FolderPath + OpeningBalanceINGPL))
            {
                if (Utils.IsOpen(FolderPath + OpeningBalanceINGPL)) openedFiles.Append($"{OpeningBalanceINGPL}\n");
            }
            else
                missedFiles.Append($"{OpeningBalanceINGPL}\n");


            if (File.Exists(FolderPath + SplitPaymentINGPL))
            {
                if (Utils.IsOpen(FolderPath + SplitPaymentINGPL)) openedFiles.Append($"{SplitPaymentINGPL}\n");
            }
            else
                missedFiles.Append($"{SplitPaymentINGPL}\n");

            if (File.Exists(FolderPath + DailyLimitINGPLEUR))
            {
                if (Utils.IsOpen(FolderPath + DailyLimitINGPLEUR)) openedFiles.Append($"{DailyLimitINGPLEUR}\n");
            }
            else
                missedFiles.Append($"{DailyLimitINGPLEUR}\n");

            if (File.Exists(FolderPath + DailyLimitINGPLPLN))
            {
                if (Utils.IsOpen(FolderPath + DailyLimitINGPLPLN)) openedFiles.Append($"{DailyLimitINGPLPLN}\n");
            }
            else
                missedFiles.Append($"{DailyLimitINGPLPLN}\n");

            // KK nowy plik z danymi SEB wg pomys³u Ewy Jeziorak
            if (File.Exists(FolderPath + OpeningBalanceSEB))
            {
                if (Utils.IsOpen(FolderPath + OpeningBalanceSEB)) openedFiles.Append($"{OpeningBalanceSEB}\n");
            }
            else
                missedFiles.Append($"{OpeningBalanceSEB}\n");

            // KK plik z danymi Santander
            if (File.Exists(FolderPath + OpeningBalanceSantander))
            {
                if (Utils.IsOpen(FolderPath + OpeningBalanceSantander)) openedFiles.Append($"{OpeningBalanceSantander}\n");
            }
            else
                missedFiles.Append($"{OpeningBalanceSantander}\n");

            // KK plik z danymi ING NL
            if (File.Exists(FolderPath + OpeningBalanceINGNL))
            {
                if (Utils.IsOpen(FolderPath + OpeningBalanceINGNL)) openedFiles.Append($"{OpeningBalanceINGNL}\n");
            }
            else
                missedFiles.Append($"{OpeningBalanceINGNL}\n");


            //sprawdzenie, czy brak plików
            if (missedFiles.ToString() != string.Empty)
            {
                errMessage.Append($"Required files listed below are not present in ROBOT folder:\n{missedFiles.ToString()}\n");
            }

            if (openedFiles.ToString() != string.Empty)
            {
                errMessage.Append($"Files listed below are opened, please close them before starting ROBOT:\n{openedFiles.ToString()}\n");
            }

            if (!String.IsNullOrEmpty(errMessage.ToString()))
            {
                errMessage.Insert(0, "Process aborted! \n\n");
                errMessage.Append("Press any key to continue.");
                message = errMessage.ToString();
                return false;
            }

            message = "Success!";
            return true;
        }

        /// <summary>
        /// W pliku z raportem generuje arkusze testowe, dla których property s¹ ustawione na true.
        /// </summary>
        /// <returns></returns>
        public static void GenerateLogSheets()
        {
            if (SplitTest)
            {
                if (logWorkbook == null) logWorkbook = GetWorkbook("logs.xlsx");
                if (logWorkbook.Workbook.Worksheets["split"] != null) logWorkbook.Workbook.Worksheets.Delete("split");
                splitLogs = logWorkbook.Workbook.Worksheets.Add("split");
                splitLogs.Cells[1, 1].Value = "Row in CF";
                splitLogs.Cells[1, 2].Value = "Description";
                splitLogs.Cells[1, 3].Value = "Desc cont VAT?";
                splitLogs.Cells[1, 4].Value = "Is Split";
                splitLogs.Cells[1, 5].Value = "Gross";
                splitLogs.Cells[1, 6].Value = "VAT";
                splitLogs.Cells[1, 7].Value = "Net";
                splitLogs.Cells[1, 8].Value = "Gr==Vat?";
                splitLogs.Cells[1, 9].Value = "OB";
                splitLogs.Cells[1, 10].Value = "Split acc";
                splitLogs.Cells[1, 11].Value = "Split";
                splitLogs.Cells[1, 12].Value = "Split Over";
                logWorkbook.Save();
            }
        }
    }
}