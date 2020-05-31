using System;
using System.Collections.Generic;
using OfficeOpenXml;
using System.Linq;

namespace RPAExcelProject
{
    public static class OpeningBalanceINGDataProcessor
    {
        const string BankName = "ING";

        const int BookingDateColumn = 4;
        const int AccountOwnerNameColumn = 2;
        const int CurrencyColumn = 6;
        const int AmountColumn = 3;
        const int DetailsColumn = 10;
        
        const int ResultStartRow = 15;
        const int ResultCcyColumn = 2;
        const int ResultSPVColumn = 4;
        const int ResultOppeningBalanceColumn = 5;



        /// <summary>
        /// lklkl
        /// </summary>
        /// <param name="bookingDate">jjkjljk</param>
        /// <returns>OLA</returns>
        static List<OpeningBalanceINGData> GetData(string bookingDate, ExcelWorksheet srcSheet)
        {
            var result = new List<OpeningBalanceINGData>();
            for (int row = srcSheet.Dimension.Start.Row + 1; row <= srcSheet.Dimension.End.Row; row++)
            {
                if (srcSheet.Cells[row, BookingDateColumn].Value != null
                    && srcSheet.Cells[row, DetailsColumn].Value.ToString().Contains("Transfer odwrotny")
                    )
                {
                    var _bookingDate = DateTime.Parse(srcSheet.Cells[row, BookingDateColumn].Value.ToString());

                    if (_bookingDate.ToString("dd.MM.yyyy") == bookingDate)
                    {
                        var data = new OpeningBalanceINGData();
                        data.AccountOwnerName = srcSheet.Cells[row, AccountOwnerNameColumn].Value.ToString().ToUpper();
                        data.BookingDate = _bookingDate.ToString("dd.MM.yyyy");
                        data.Currency = srcSheet.Cells[row, CurrencyColumn].Value?.ToString();
                        double _amount;
                        var _tmpAmount = srcSheet.Cells[row, AmountColumn].Value;

                        if (_tmpAmount != null)
                        {
                            Double.TryParse(_tmpAmount.ToString(), out _amount);
                            data.Amount = _amount;
                        }
                        else
                            data.Amount = 0.0;

                        result.Add(data);
                    }
                }
            }

            return result;
        }

        public static void INGProcessData(string bookingDate, ExcelWorksheet destSheet, ExcelWorksheet srcSheet)
        {
            var data = GetData(bookingDate, srcSheet);
            for (int row = ResultStartRow; row <= destSheet.Dimension.End.Row; row++)
            {
                var company = destSheet.Cells[row, ResultSPVColumn].Value?.ToString().ToUpper();
                var valueColumn = destSheet.Cells[row, ResultCcyColumn].Value?.ToString();
                string currency = "";
                string bank = "";

                if (valueColumn != null)
                {
                    valueColumn = valueColumn.Replace(" ", "");
                    currency = valueColumn.Substring(1 - 1, 3);
                    bank = valueColumn.Substring(4 - 1, 3);
                }

                var result = data.Where(x => x.AccountOwnerName == company && x.Currency == currency).FirstOrDefault();

                if (result != null)
                    if (result.Amount != 0 && bank == BankName)
                        destSheet.Cells[row, ResultOppeningBalanceColumn].Value = result.Amount;
            }

        }
    }
}
