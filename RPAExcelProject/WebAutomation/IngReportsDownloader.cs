using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace RPAExcelProject
{
    public class IngReportsDownloader : IDisposable
    {
        private Browser _browser;
        private ChromeDriver _driver;

        public void Login()
        {
            _browser = new Browser("Ing");
            _browser.Start();
            _driver = _browser.Driver;

            LoginToIng();
        }

        public void Execute()
        {
            _driver.SwitchTo().DefaultContent();
            _driver.SwitchTo().Frame(0);

            try
            {

                var filesByCompany = new Dictionary<string, string[]>() {
                            { "Skanska Spółka Akcyjna", new[] { "LIMITY DZIENNE_LIMITY DZIENNE PLN_", "LIMITY DZIENNE_Limity EUR_" } },
                            { "Towarzystwo Funduszy Inwestycyjnych", new[] { "TFI SALDA_", "TFI_flows_new_TFI_all_" } },
                            { "SKANSKA PROPERTY POLAND FIZ AKTYWÓW NIEPUBLICZNYCH", new[] { "FIZAN_salda_", "SPP fizan_FIZAN_daily flows_" } },
                            { "SFS", new[] { "SFS All_VAT_", "SALDA SPLIT_",  "Daily Flows Poland_ROBOT 2_"} }
                        };

                foreach (var company in filesByCompany.Keys)
                {
                    _browser.Invoke($"Company context {company}", () => ChangeCompanyContext(company));

                    _browser.Invoke($"Go to import export {company}", () =>
                    {
                        if (GoToImportExport() == false)
                        {
                            GoToImportExport();
                        };
                    });


                    foreach (var file in filesByCompany[company])
                    {
                        _browser.Invoke($"Go to exported files {company}", () => GoToExportedFiles(file));
                    }

                }

            }
            catch (Exception ex)
            {
                _browser.TakeScreenshot("Failed.png");
                _browser.Log(ex.ToString());
            }

        }

        private void GoToExportedFiles(string filePrefix)
        {
            Thread.Sleep(2000);

            _driver.FindElement(By.LinkText("Pliki wyeksportowane")).Click();

            _browser.WaitUntil(By.CssSelector("#MOD_IEX_EFL_table > div > div.csf-grid-body.ng-scope > div > div:nth-child(1) > div.csf-grid-row-container > div > div > div.csf-grid-table-cell.table-column-one.word-break-all.ng-scope > a"));

            _driver.FindElement(By.Id("_acc_input")).Clear();
            _driver.FindElement(By.Id("_acc_input")).SendKeys(filePrefix);
            _driver.FindElement(By.Id("_search_button")).Click();

            _browser.WaitUntil(By.CssSelector("#MOD_IEX_EFL_table > div > div.csf-grid-body.ng-scope > div > div:nth-child(1) > div.csf-grid-row-container > div > div > div.csf-grid-table-cell.table-column-one.word-break-all.ng-scope > a"));
            Thread.Sleep(3000);

            var table = _driver.FindElement(By.ClassName("exported-files-list-table"));
            if (table != null)
            {
                // TakeScreenshot(log, _driver, $"plikiwyeksportowane_{filePrefix}.png");
                var items = table.FindElements(By.TagName("a"));


                _browser.Log($"Links for {filePrefix}");
                foreach (var item in items)
                {
                    _browser.Log(item.Text);
                }

                if (items.First()?.Text?.StartsWith(filePrefix) == true)
                {
                    // new Actions(driver).MoveToElement(items[3]).Perform();                    
                    // Thread.Sleep(500);
                    _browser.TakeScreenshot($"Before_click_{items.First().Text}.png");
                    items.First()?.Click();
                    Thread.Sleep(500);
                }
            }
        }

        private void ChangeCompanyContext(string companyName)
        {
            new Actions(_driver).MoveToElement(_driver.FindElement(By.CssSelector("#_selectize_input"))).Perform();

            _driver.FindElement(By.CssSelector("#MOD_DSH_HDR_owner-name .company-ellipsis")).Click();
            _driver.FindElement(By.Id("_INP_TXT")).SendKeys(companyName);
            _browser.TakeScreenshot($"company_afterinput_{companyName}.png");
            _driver.FindElement(By.Id("_INP_TXT")).SendKeys(Keys.Enter);
            _browser.Log("enter");

            _browser.WaitUntil(By.ClassName("nav-menu-configuration"));

            Thread.Sleep(15000);
            //TakeScreenshot(log, _driver, "company_afterinput2.png");

            var input = _driver.FindElement(By.CssSelector("#_selectize_input > div"));
            if (input != null)
            {
                _browser.Log($"Found input: {input.Text}");

            }
        }

        private bool GoToImportExport()
        {
            try
            {
                var navMenu = _driver.FindElement(By.ClassName("nav-menu-configuration"));
                if (navMenu != null)
                {
                    _browser.Log("Menu found");
                }

                Actions hover = new Actions(_driver);
                hover.MoveToElement(navMenu).Perform();
                //TakeScreenshot(log, _driver, "HoverMenu.png");

                hover = new Actions(_driver);
                var admMenu = _driver.FindElement(By.CssSelector("#MOD_DSH_HDR_navmenu-right-configurations > span"));
                if (admMenu != null)
                {
                    _browser.Log("Adm Menu found");
                }

                hover.MoveToElement(admMenu).Perform();
                //TakeScreenshot(log, _driver, "HoverMenu_Adm.png");

                var importExportMenu = _driver.FindElement(By.Id("MOD_DSH_HDR_navmenu-right-configurations-import-export"));


                Thread.Sleep(3000);
                importExportMenu.Click();
                _browser.WaitUntil(By.ClassName("import-export-container"));
            }
            catch (Exception)
            {
                _browser.TakeScreenshot("importExport.png");
                _browser.Log("Failed to get to importExport");
                return false;
            }

            return true;

        }


        private void LoginToIng()
        {
            _driver.Navigate().GoToUrl("https://start.ingbusiness.pl/ing2/login/");
            _driver.FindElement(By.Id("alias-input")).Click();
            _browser.Log("Waiting to login");

            var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(120));

            wait.Until(x =>
            {
                string url = "https://start.ingbusiness.pl/ing2/index.jsp";
                return x.Url.Equals(url, StringComparison.OrdinalIgnoreCase);
            });

            Thread.Sleep(500);

            _browser.WaitUntil(By.TagName("frameset"));

            Thread.Sleep(5000);
            //TakeScreenshot(log, _driver, "AfterLogin.png");
        }

        public void Dispose()
        {
            _browser?.Dispose();
        }
    }





}

