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
        private string excelPath;
        private int flows;
        private int lenghtOfFlow;
        private List<List<Elem>> output;
        /// <summary>
        /// Пустой конструктор по умолчанию
        /// </summary>
        public Parser() { }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="driverPath">путь к папке с драйвером</param>
        /// <param name="excelPath">путь до файла с кад. номерами</param>
        /// <param name="flows">число потоков</param>
        /// <param name="lenghtOfFlow">количество элементов в одном потоке</param>
        public Parser(string driverPath, string excelPath, int flows, int lenghtOfFlow)
        {
            this.driverPath = driverPath;
            this.excelPath = excelPath;
            this.flows = flows;
            this.lenghtOfFlow = lenghtOfFlow;
            output = new List<List<Elem>>();
        }

        
        /// <summary>
        /// Многопоточный запуск парсера
        /// </summary>       
        private void multyTask()
        {
            for (int i = 0; i < output.Count; i += flows)
            {
                // Пока число необработанный элементов больше количеств потоков
                if (output.Count - i >= flows)
                {
                    Task[] tasks1 = new Task[flows];
                    for (var j = 0; j < tasks1.Length; j++)
                    {
                        var index = i + j;
                        tasks1[j] = Task.Factory.StartNew(() => { parser(output[index]); });
                    }
                    Task.WaitAll(tasks1); // ожидаем завершения задач 
                }
                // Создаем потоки на оставшееся число элементов output
                else
                {
                    int N = output.Count - i;

                    Task[] tasks2 = new Task[N];
                    for (var j = 0; j < tasks2.Length; j++)
                    {
                        var index = i + j;
                        tasks2[j] = Task.Factory.StartNew(() => { parser(output[index]); });
                    }
                    Task.WaitAll(tasks2);
                }
            }
        }
        
        /// <summary>
        /// Выполняет парсинг и запись результата в коллекцию output
        /// </summary>
        /// <param name="elem">Элемент коллекции output</param>
        /// <value name="counter">Каунтер для цикла while</value>
        /// <value name="cntr">Каунтер для цикла тупняка</value>
        /// <value name="fr">Фрейм поиска</value>
        private void parser(List<Elem> elem)
        {
            int cnt = 0;
            #region Driver initializtion
            // Инициализация драйвера
            IWebDriver driver = new ChromeDriver(driverPath);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            // URL с открытым полем поиска. Не менять!
            driver.Url = @"https://pkk5.rosreestr.ru/#x=1770771.834433252&y=10055441.599232893&z=3&app=search&opened=1";
            #endregion
            Thread.Sleep(1000);
            #region Choose OKS search
            // Выбор поиска по объектам ОКС
            IWebElement fr = driver.FindElement(By.CssSelector(@"#app-search-form > div > div > div > div > button"));
            Thread.Sleep(1000);
            fr.Click();
            fr = driver.FindElement(By.CssSelector(@"#tag_5"));
            fr.Click();
            #endregion

            for (int i = 0; i < elem.Count; i++)
            {
                string cad_num = elem[i].cad_num;
                try
                {
                   
                    int counter = 0;                    
                    int cntr = 0;

                    #region Выбор фрейма поиска, очистка, запись кадастрового номера и нажатие на кнопку поиска
                    fr = driver.FindElement(By.CssSelector(@"#search-text"));
                    fr.Clear();
                    fr.SendKeys(cad_num);
                    Thread.Sleep(500);
                    driver.FindElement(By.CssSelector(@"#app-search-submit")).Click();
                    #endregion

                    // Ожидание подгрузки элемента DOM. Таймаут - 10 сек.
                    //TODO: Сделать проверку на отсутствие кдастрового номера в базе
                    wait.Until(ExpectedConditions.ElementExists(By.CssSelector(@"#feature-oks-info > div")));

                    // Цикл ожидания подгрузки элемента DOM. Цикл был реализован до включения в код wait.Until
                    // TODO: Проверить его необхдоимость, при наличии wait.Until                    
                    while (counter < 20)
                    {
                        Thread.Sleep(500);
                        // pane - панель с результатами поиска
                        var pane = driver.FindElements(By.CssSelector(@"#feature-oks-info > div"));
                        // Спаршенная строка. Вся смысловая нагрузка выделена с помощью #
                        var val = Regex.Replace(pane[0].Text, @"(\r\n)", "#", RegexOptions.Compiled);
                        // condition - текст в поле "Тип". Служит для проверки подгрузки DOM.
                        var condition = Regex.Match(val, @"Тип:#([^#]+)#", RegexOptions.Compiled).Groups[1].Value;
                        // Для проверки тупняка. Когда старая панель висит и успешно парсится
                        // а панель под новый кадастровый номер не прогружается. 
                        var cad_numFromPane = Regex.Match(val, @"Кад. номер:#([^#]+)#", RegexOptions.Compiled).Groups[1].Value;
                        var equal = cad_num.Equals(cad_numFromPane);

                        // Проверка на тупняк. Не работает. Можно снижать каунтер проверки.
                        // TODO: Сменить логику. Попробовать остановку прогрузки страницы.
                        while (!equal)
                        {
                            Thread.Sleep(500);
                            pane = driver.FindElements(By.CssSelector(@"#feature-oks-info > div"));
                            cad_numFromPane = Regex.Match(val, @"Кад. номер:#([^#]+)#", RegexOptions.Compiled).Groups[1].Value;
                            cntr++;
                            if (cntr > 50) break;
                            Console.WriteLine("тупняк " + cntr);                            
                        }
                        cntr = 0;

                        // лучший случай. DOM загрузился и кад. номера идентичны
                        if (!condition.Equals("-") && equal)
                        {
                            OKS oks = new OKS(val, cad_num);
                            elem[i].oks = oks;

                            break;
                        }
                        // сомнительный случай
                        // TODO: Проверить его достижимость
                        else if (counter == 19)
                        {
                            OKS oks = new OKS(val, cad_num);
                            elem[i].oks = oks;

                            counter++;
                        }
                        else
                        {
                            counter++;
                        }
                    }

                    Thread.Sleep(500);
                    cnt++;
                    //Console.WriteLine(cnt + ") For CAD_NUM " + cad_num + "    counter = " + counter);
                }
                catch (Exception e)
                {
                    // Эти два исключения должны уйти, когда будет включена проверка на отсутствие результата поиска
                    var name = e.GetType().Name;
                    if (name.Equals("ArgumentOutOfRangeException") ||
                        name.Equals("WebDriverTimeoutException"))
                    {
                        OKS oks = new OKS(cad_num, "cad_num doesn't exist", -999);
                        elem[i].oks = oks;
                    }
                    else
                    {
                        OKS oks = new OKS(cad_num, name, -999);
                        elem[i].oks = oks;
                    }
                }
            }
            driver.Close();           
        }

        private List<Elem> reSearchParser(List<string> reSearch)
        {
            var result = new List<Elem>();

            //string path = @"C:\Users\vtsvetkov\source\repos\pkk_5_parser";
            IWebDriver driver = new ChromeDriver(driverPath);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            driver.Url = @"https://pkk5.rosreestr.ru/#x=1770771.834433252&y=10055441.599232893&z=3&app=search&opened=1";

            Thread.Sleep(1000);

            IWebElement fr = driver.FindElement(By.CssSelector(@"#app-search-form > div > div > div > div > button"));
            Thread.Sleep(1000);
            fr.Click();
            fr = driver.FindElement(By.CssSelector(@"#tag_5"));
            fr.Click();

            int countOfTry = 0;

            for (int i = 0; i < reSearch.Count; i++)
            {
                string cad_num = reSearch[i];
                try
                {
                    int counter = 0;
                    int cntr = 0;

                    fr = driver.FindElement(By.CssSelector(@"#search-text"));
                    fr.Clear();
                    fr.SendKeys(cad_num);

                    Thread.Sleep(1000);

                    driver.FindElement(By.CssSelector(@"#app-search-submit")).Click();

                    wait.Until(ExpectedConditions.ElementExists(By.CssSelector(@"#feature-oks-info > div")));

                    while (counter < 20)
                    {
                        Thread.Sleep(500);
                        var pane = driver.FindElements(By.CssSelector(@"#feature-oks-info > div"));
                        var val = Regex.Replace(pane[0].Text, @"(\r\n)", "#", RegexOptions.Compiled);
                        var condition = Regex.Match(val, @"Тип:#([^#]+)#", RegexOptions.Compiled).Groups[1].Value;
                        var cad_numFromPane = Regex.Match(val, @"Кад. номер:#([^#]+)#", RegexOptions.Compiled).Groups[1].Value;
                        var equal = cad_num.Equals(cad_numFromPane);

                        while (!equal)
                        {
                            Thread.Sleep(500);
                            pane = driver.FindElements(By.CssSelector(@"#feature-oks-info > div"));
                            cad_numFromPane = Regex.Match(val, @"Кад. номер:#([^#]+)#", RegexOptions.Compiled).Groups[1].Value;
                            cntr++;
                            if (cntr > 50) break;
                            Console.WriteLine("тупняк " + cntr);
                        }
                        cntr = 0;


                        if (!condition.Equals("-") && equal)
                        {
                            OKS oks = new OKS(val, cad_num);
                            result.Add(new Elem(reSearch[i], oks));
                            countOfTry = 0;

                            break;
                        }
                        else
                        {
                            counter++;
                        }
                    }
                    Thread.Sleep(500);
                }
                catch (Exception e)
                {
                    if (countOfTry < 5)
                    {
                        countOfTry++;
                        i--;
                    }
                    else
                    {
                        countOfTry = 0;
                        var name = e.GetType().Name;
                        if (name.Equals("ArgumentOutOfRangeException") ||
                            name.Equals("WebDriverTimeoutException"))
                        {
                            OKS oks = new OKS(cad_num, "cad_num doesn't exist", -999);
                            result[i].oks = oks;
                        }
                        else
                        {
                            OKS oks = new OKS(cad_num, e.GetType().Name, -999);
                            result[i].oks = oks;
                        }
                    }
                }
            }
            driver.Close();
            return result;
        }

        private void WebDriverInitSearchOKS()
        {     
                    
            IWebDriver driver = new ChromeDriver(driverPath);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            
            driver.Url = @"https://pkk5.rosreestr.ru/#x=1770771.834433252&y=10055441.599232893&z=3&app=search&opened=1";
            
            Thread.Sleep(1000);
            
            
            IWebElement fr = driver.FindElement(By.CssSelector(@"#app-search-form > div > div > div > div > button"));
            Thread.Sleep(1000);
            fr.Click();
            fr = driver.FindElement(By.CssSelector(@"#tag_5"));
            fr.Click();
            
        }

        public void RunParsingOKS()
        {
            //ExcelApp();
            multyTask();
            //createResultExcel();
        }        
    }
}
