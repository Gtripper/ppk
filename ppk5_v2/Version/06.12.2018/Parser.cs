using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;



namespace ppk5_v2
{
    /// <summary>
    /// Парсер росрреста (ppk5.rosreestr.ru) по кадастровым номерам
    /// </summary>    
    public class Parser : IParser
    {
        private string driverPath;
        private List<Elem> elem;
        
        public Parser() { }

        public Parser(string driverPath, List<Elem> elem)
        {
            this.driverPath = driverPath;
            this.elem = elem;
        }

        /// <summary>
        /// Выполняет парсинг и запись результата в коллекцию output
        /// </summary>
        /// <param name="elem">Элемент коллекции output</param>
        /// <value name="counter">Каунтер для цикла while</value>
        /// <value name="cntr">Каунтер для цикла тупняка</value>
        /// <value name="fr">Фрейм поиска</value>
        public void parser()
        {
            var driver = WebDriverInitSearchOKS();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            WebDriverWait waitMs = new WebDriverWait(driver, TimeSpan.FromMilliseconds(500));

            foreach (var val in elem)
            {
                string cad_num = val.cad_num;
                try
                {
                    InputTextToSearchBox(driver, cad_num);

                    if (NoResult(driver))
                    {
                        OKS oks = new OKS(cad_num, "cad_num doesn't exist", 999);
                        val.oks = oks;
                    }
                    else
                    {
                        

                        var parsedString = PaneOKS(driver, wait);

                        var cad_numFromPane = Regex.Match(parsedString, @"Кад. номер:#([^#]+)#", RegexOptions.Compiled).Groups[1].Value;                        
                        var equal = cad_num.Equals(cad_numFromPane);

                        if (equal)
                        {
                            OKS oks = new OKS(parsedString, cad_num);
                            val.oks = oks;
                        }
                        Thread.Sleep(500);                        
                    }
                }
                catch (Exception e)
                {
                    // Эти два исключения должны уйти, когда будет включена проверка на отсутствие результата поиска
                    var name = e.GetType().Name;
                    if (name.Equals("ArgumentOutOfRangeException") ||
                        name.Equals("WebDriverTimeoutException"))
                    {
                        OKS oks = new OKS(cad_num, "cad_num doesn't exist", -999);
                        val.oks = oks;
                    }
                    else
                    {
                        OKS oks = new OKS(cad_num, name, -999);
                        val.oks = oks;
                        Console.WriteLine(e.StackTrace);
                        Console.WriteLine(cad_num + "   " + name);
                    }
                }
            }
            driver.Close();
        }        

        private IWebDriver WebDriverInitSearchOKS()
        {
            try
            {
                IWebDriver driver = new ChromeDriver(driverPath);

                driver.Url = @"https://pkk5.rosreestr.ru/#x=1770771.834433252&y=10055441.599232893&z=3&app=search&opened=1";

                Thread.Sleep(1000);

                IWebElement fr = driver.FindElement(By.CssSelector(@"#app-search-form > div > div > div > div > button"));
                Thread.Sleep(500);
                fr.Click();
                fr = driver.FindElement(By.CssSelector(@"#tag_5"));
                fr.Click();

                return driver;
            }
            catch
            {
                ///TODO: Добавить журналирование в еррорлог
                WebDriverInitSearchOKS();
                return null;
            }
        }

        private void InputTextToSearchBox(IWebDriver driver, string text)
        {
            Thread.Sleep(500);
            var fr = driver.FindElement(By.CssSelector(@"#search-text"));
            fr.Clear();
            fr.SendKeys(text);
            Thread.Sleep(500);
            driver.FindElement(By.CssSelector(@"#app-search-submit")).Click();
        }

        private bool NoResult(IWebDriver driver)
        {
            try
            {
                Thread.Sleep(500);
                var noResult = driver.FindElement(By.CssSelector(@"#searchEngineFeatureSet_list > div > div"));
                return noResult.Displayed;
            }
            catch
            {
                return false;
            }
        }

        private string PaneOKS(IWebDriver driver, WebDriverWait wait)
        {
            var result = "";
            try
            {
                wait.Until(p => !p.FindElement(By.CssSelector
                            (@"#feature-oks-info > div > div:nth-child(1) > div.col-xs-8.col-lg-8.col-sm-8.col-md-8"))
                                .Text.Equals("-"));                
            }
            catch (StaleElementReferenceException e)
            {
                Thread.Sleep(500);
                Console.WriteLine(e);
                PaneOKS(driver, wait);
            }
            var pane = driver.FindElements(By.CssSelector(@"#feature-oks-info > div"));
            result = Regex.Replace(pane[0].Text, @"(\r\n)", "#", RegexOptions.Compiled);
            return result;
        }

        public void RunParsingOKS()
        {
            //ExcelApp();
            //multyTask();
            //createResultExcel();
        }
    }
}
