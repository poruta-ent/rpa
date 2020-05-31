using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;

namespace RPAExcelProject
{
    public class Browser : IDisposable
    {
        private StreamWriter _log;

        public ChromeDriver Driver { get; private set; }

        public Browser(string description)
        {
            _log = File.CreateText($"Log_{description}.txt");
        }

        public void Start()
        {
            var chromeOptions = new ChromeOptions();
            chromeOptions.AddArguments(
                "--no-default-browser-check",
                "--disable-extensions",
                "--disable-popup-blocking",
                "--ignore-certificate-errors"
               );


            chromeOptions.AddUserProfilePreference("credentials_enable_service", false);
            chromeOptions.AddUserProfilePreference("download.default_directory", Environment.CurrentDirectory);
            chromeOptions.AddUserProfilePreference("download.prompt_for_download", false);
            chromeOptions.AddUserProfilePreference("download.directory_upgrade", true);
            chromeOptions.AddUserProfilePreference("safebrowsing.enabled", false);
            chromeOptions.AddUserProfilePreference("safebrowsing.disable_download_protection", true);

            this.Driver = new ChromeDriver(Environment.CurrentDirectory, chromeOptions);
            this.Driver.Manage().Window.Maximize();
        }

        public void TakeScreenshot(string name)
        {
            var screenshot = Driver.GetScreenshot();
            screenshot.SaveAsFile(name, ScreenshotImageFormat.Png);
            Log($"Saved screenshot: {name}");
        }

        public void Log(string msg)
        {
            _log.WriteLine($"{DateTime.Now.ToString("HH:mm:ss.fff")} {msg}");
            _log.Flush();
        }

        public void Invoke(string description, Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                TakeScreenshot($"Failed_{description}.png");
                Log(ex.ToString());
            }
        }
        
        public void WaitUntil(By findBy, int howLong = 15 )
        {
            var wait = new WebDriverWait(Driver, TimeSpan.FromSeconds(howLong));
            wait.Until(x => x.FindElement(findBy));
        }

        public void Dispose()
        {
            Driver?.Dispose();
            _log?.Dispose();
        }
    }


   

    
}

