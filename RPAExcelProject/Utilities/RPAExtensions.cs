using OfficeOpenXml;

namespace RPAExcelProject
{
    public static class RPAExtensions
    {
        public static string GetNotNullString(this ExcelRange cell)
        {
            return cell.Value != null ? cell.Value.ToString() : string.Empty;
        }

        public static float GetNotNullFloat(this ExcelRange cell)
        {
            if (float.TryParse(cell.GetNotNullString(), out float notNullFloat))
            {
                return notNullFloat;
            }
            else
            {
                return 0f;
            }
        }
    }
}
