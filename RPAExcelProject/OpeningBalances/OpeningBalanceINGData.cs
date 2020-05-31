
namespace RPAExcelProject
{
    /// <summary>
    /// Struktura do odczytania z pliku OpeningBalanceINGPL.xlsx danych z danego dnia
    /// </summary>
    public class OpeningBalanceINGData
    {
        public string AccountOwnerName { get; set; }
        public string BookingDate { get; set; }
        public string Currency { get; set; }
        public double Amount { get; set; }
    }
}
