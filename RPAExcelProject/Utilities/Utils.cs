using System;
using OfficeOpenXml;
using System.IO;

namespace RPAExcelProject
{
    public static class Utils
    {

        public static bool CompareStringDates(string date, string dateToCompareWith)
        {
            return (DateTime.TryParse(date, out DateTime dateToCompare) && dateToCompare.ToString("dd.MM.yyyy") == dateToCompareWith);
        }


        //Sprawdza czy plik jest otwarty
        public static bool IsOpen(string filePath)
        {
            try
            {
                FileStream fs = File.Open(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None);
                fs.Close();
            }
            catch (IOException)
            {
                return true;
            }
            return false;
        }
        
        //Zwraca informację czy arkusz z daną nazwą w danej dacie już istnieje w danym pliku
        public static bool CheckIfSheetExists(ExcelPackage reportWorkbook, string sheetName)
        {

            if (reportWorkbook.Workbook.Worksheets[sheetName] != null)
            {
                return true;
            }
            
            return false;
        }

        //wyświetla komunikat - póki co w konsoli
        public static void ShowUserMessage(string errMessage)
        {
            Console.WriteLine(errMessage);
        }

        // KK ile dni odjąć od daty, żeby nowa data była w zakresie Poniedziałku do Piątku w poprzednim tygodniu
        public static int LotOfDays(DayOfWeek _dayOfWeek)
        {
            int _days = 0;
            if ((int)_dayOfWeek >= 2 && (int)_dayOfWeek <= 6)
            {
                _days = -1;
            }
            else if ((int)_dayOfWeek == 1)
            {
                _days = -3;
            }
            else if ((int)_dayOfWeek == 0)
            {
                _days = -2;
            }
            return _days;
        }

        // KK metoda, nie uwzględnia dni świątecznych i innych wolnych od pracy
        /// <summary>
        /// Podaje datę poprzednią w odniesieniu do daty z parametru;
        /// data taka mieści się w zakresie Poniedziałek do Piątek;
        /// dla daty oznaczającej Poniedziałek jakiegoś tygodnia zwracana jest data z Piątku poprzedniego tygodnia
        /// </summary>
        /// <param name="_date"></param>
        /// <returns></returns>
        public static DateTime PreviousDate(DateTime _date)
        {
            int _days = LotOfDays(_date.DayOfWeek);

            return _date.AddDays(_days);
        }

        // KK odczytanie pliku
        public static string CsvFileName(string _bankFileMask)
        {
            //string _plik = "Informacje o saldach*" + _dataZn + "*.csv";

            string[] _files = Directory.GetFiles(ReportFile.FolderPath, _bankFileMask, SearchOption.TopDirectoryOnly);
            if (_files.Length > 0)
            {
                string _file = _files[0];
                var kawalki = _file.Split("\\");
                return kawalki[kawalki.Length - 1];
            }

            return "??";
        }

        //KK pozyskanie plików .csv
        /// <summary>
        /// Znajduje niezbędne pliki .csv; parametr: data poprzedni dzień niż data raportu
        /// </summary>
        /// <param name="dateTime"></param>
        //public static void FilesFromCsv(DateTime dateTime)
        public static void FilesFromCsv()
        {
            //bank SEB, w nazwie pliku data z dnia poprzedniego
            string _fileBaseName = "SEB_ROBOT_NEW ";
            string _fileExtention = ".csv";
            string _fileDate = MasterData.previousReportDate.ToString("yyyyMMdd");

            ReportFile.OpeningBalanceSEB = _fileBaseName + _fileDate + _fileExtention;

            //bank Santander, w nazwie pliku data z dnia raportu
            _fileDate = DateTime.Parse(MasterData.reportDate).ToString("yyyyMMdd");
            string _fileMask = "Informacje o saldach*" + _fileDate + "*" + _fileExtention;

            ReportFile.OpeningBalanceSantander = Utils.CsvFileName(_fileMask);
            if (ReportFile.OpeningBalanceSantander == "??")
                ReportFile.OpeningBalanceSantander = _fileMask;

            //bank ING NL dane do ING BV, w nazwie pliku data z dnia raportu
            _fileMask = "WB*Balances*" + _fileDate + _fileExtention;

            ReportFile.OpeningBalanceINGNL = Utils.CsvFileName(_fileMask);
            if (ReportFile.OpeningBalanceINGNL == "??")
                ReportFile.OpeningBalanceINGNL = _fileMask;

        }

    }
}
