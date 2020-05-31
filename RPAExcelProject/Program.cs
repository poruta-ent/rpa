using System;
using System.Text;

namespace RPAExcelProject
{
    class Program
    {
        static void Main(string[] args)
        {


            if(args.Length >0 && args[0] == "download")
            {
                new ReportDownloader().Start();
            }
            else
            {
                //KK dzięki Maćkowi Szczepańskiemu poniższa komenda do zarejestrowania "klubu" sposbów kodowania znaków
                //dzięki temu można wykorzystać stronę kodową 1250 jako znaki polskie w Excel przy załadowaniu pliku .csv 
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                RobotManager.RunCfRobot();
                Console.ReadKey();
            }
            
           
        }
    }
}
