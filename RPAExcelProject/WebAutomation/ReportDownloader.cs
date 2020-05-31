using SeleniumExtras.WaitHelpers;
using System.Text;

namespace RPAExcelProject
{

    public class ReportDownloader
    {
        public void Start()
        {
            var ingDownloader = new IngReportsDownloader();

            try
            {
                ingDownloader.Login();

                ingDownloader.Execute();
            }
            finally
            {
                ingDownloader.Dispose();
            }

        }

    }


   

    
}

