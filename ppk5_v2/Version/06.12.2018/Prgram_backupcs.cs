using System;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using ppk5_v2;

namespace pkk_5_parser
{
    public class Program
    {
        static void Main(string[] args)
        {

            string path = @"C:\Users\vtsvetkov\source\repos\pkk_5_parser\pkk_5_parser\resourses\test.xlsx";
            int lenghtOfRow = 500;
            int Flows = 6;
            string driverPath = @"C:\Users\vtsvetkov\source\repos\pkk_5_parser";
            //ExcelApp(path, driverPath, lenghtOfRow, Flows);

            Console.WriteLine(Error_log.Error_TimeStamp());

            //ExcelApp excelApp = new ExcelApp(path, "A", 2);

            var list = Tests.testElem500().ToList();

            //Parser parser = new Parser(driverPath, list);
            //parser.parser();

            MultiThread multiThread = new MultiThread(list, 4, 50, driverPath);
            multiThread.ThreadMaster();

            Console.WriteLine(Error_log.Error_TimeStamp());
            var qwe = list.Where(p => p.oks.type.Contains("Exception"));
            //Console.WriteLine("WIN");

            Console.Read();
        }

        public static void ExcelApp(string path, string driverPath, int lenghtOfRow, int flowCount)
        {
            #region Open and prepare Excel File

            Application app = new Application();
            app.Visible = true;
            app.Workbooks.Open(path);
            Worksheet workSheet = app.ActiveSheet;

            workSheet.Cells[1, "A"] = "CAD_NUM_Original";
            workSheet.Cells[1, "B"] = "CAD_NUM";
            workSheet.Cells[1, "C"] = "Тип";
            workSheet.Cells[1, "D"] = "Наименование";
            workSheet.Cells[1, "E"] = "Адрес";
            workSheet.Cells[1, "F"] = "Форма собственности";
            workSheet.Cells[1, "G"] = "Общая площадь";
            workSheet.Cells[1, "H"] = "Минимальная этажность";
            workSheet.Cells[1, "I"] = "Максимальная этажность";
            workSheet.Cells[1, "J"] = "Подземная этажность";
            workSheet.Cells[1, "K"] = "Назначение";
            workSheet.Cells[1, "L"] = "Год ввода";
            workSheet.Cells[1, "O"] = "IncorrectCads";
            workSheet.Cells[1, "P"] = "Value";

            workSheet.Range["A1", "P1"].Font.Bold = 1;
            workSheet.Columns["A:P"].AutoFit();
            #endregion

            int iter = 13002;
            string cad_num = "";
            List<string> listCadNums = new List<string>();
            var output = new List<List<Elem>>();
            var temp = new List<Elem>();


            //(workSheet.Cells[iter, "A"].Value != null)
            while (workSheet.Cells[iter, "A"].Value != null)
            {
                cad_num = (string)(workSheet.Cells[iter, "A"] as Range).Value;

                listCadNums.Add(cad_num);
                iter++;

                temp.Add(new Elem(cad_num));

                if (iter % lenghtOfRow == 2 && iter > 0)
                {
                    output.Add(temp);
                    temp = new List<Elem>();
                }

                if (workSheet.Cells[iter, "A"].Value == null &&
                    temp.Count() < lenghtOfRow + 2)
                {
                    output.Add(temp);
                }

            }
            Console.WriteLine(iter);
            try
            {
                multyTask(output, flowCount, driverPath);
            }
            catch
            {
                Console.WriteLine("OOOPS TASK GET EXCEPTIONS");
            }
            createResultExcel(output, driverPath);
        }

        static void multyTask(List<List<Elem>> output, int flowCount, string driverPath)
        {
            for (int i = 0; i < output.Count; i += flowCount)
            {
                if (output.Count - i >= flowCount)
                {
                    Task[] tasks1 = new Task[flowCount];
                    for (var j = 0; j < tasks1.Length; j++)
                    {
                        var index = i + j;
                        tasks1[j] = Task.Factory.StartNew(() => { parser(output[index], driverPath); });
                    }
                    Task.WaitAll(tasks1); // ожидаем завершения задач 
                }
                else
                {
                    int N = output.Count - i;

                    Task[] tasks2 = new Task[N];
                    for (var j = 0; j < tasks2.Length; j++)
                    {
                        //new Task(() => { parser(output[i + j]); });
                        var index = i + j;
                        tasks2[j] = Task.Factory.StartNew(() => { parser(output[index], driverPath); });
                    }
                    Task.WaitAll(tasks2);
                }
            }
        }

        static void createResultExcel(List<List<Elem>> output, string driverPath)
        {
            var reSearch = new List<string>();


            foreach (var val in output)
            {
                foreach (var vall in val)
                {
                    try
                    {
                        if (vall.oks.type.Contains
                                ("Exception"))                      // Проверка на эксэпшн
                            reSearch.Add(vall.cad_num);
                        if (!vall.cad_num.Equals(vall.oks.cad_num))  // Проверка на тупняк (не равенство кадастра поиска и найденного)
                            reSearch.Add(vall.cad_num);
                    }
                    catch
                    {
                        reSearch.Add(vall.cad_num);
                    }
                }


            }
            var corrrectedCads = reSearchParser(reSearch, driverPath);
            #region Open and preapre new Excel File 
            var excelApp = new Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Add();

            Worksheet workSheet = excelApp.ActiveSheet;

            workSheet.Cells[1, "A"] = "CAD_NUM_Original";
            workSheet.Cells[1, "B"] = "CAD_NUM";
            workSheet.Cells[1, "C"] = "Тип";
            workSheet.Cells[1, "D"] = "Наименование";
            workSheet.Cells[1, "E"] = "Адрес";
            workSheet.Cells[1, "F"] = "Форма собственности";
            workSheet.Cells[1, "G"] = "Общая площадь";
            workSheet.Cells[1, "H"] = "Минимальная этажность";
            workSheet.Cells[1, "I"] = "Максимальная этажность";
            workSheet.Cells[1, "J"] = "Подземная этажность";
            workSheet.Cells[1, "K"] = "Назначение";
            workSheet.Cells[1, "L"] = "Год ввода";
            workSheet.Cells[1, "O"] = "IncorrectCads";
            workSheet.Cells[1, "P"] = "Value";

            workSheet.Range["A1", "P1"].Font.Bold = 1;
            workSheet.Columns["A:P"].AutoFit();
            #endregion

            var iter = 2;
            foreach (var val in output)
            {
                foreach (var vall in val)
                {
                    if (!reSearch.Contains(vall.cad_num))
                    {
                        #region Write OKS
                        workSheet.Cells[iter, "A"] = vall.cad_num;
                        workSheet.Cells[iter, "B"] = vall.oks.cad_num;
                        workSheet.Cells[iter, "C"] = vall.oks.type;
                        workSheet.Cells[iter, "D"] = vall.oks.name;
                        workSheet.Cells[iter, "E"] = vall.oks.adress;
                        workSheet.Cells[iter, "F"] = vall.oks.ownership;
                        workSheet.Cells[iter, "G"] = vall.oks.summaryArea;
                        workSheet.Cells[iter, "H"] = vall.oks.minFloors;
                        workSheet.Cells[iter, "I"] = vall.oks.maxFloors;
                        workSheet.Cells[iter, "J"] = vall.oks.numsOfUndergroundFloor;
                        workSheet.Cells[iter, "K"] = vall.oks.function;
                        workSheet.Cells[iter, "L"] = vall.oks.years;
                        workSheet.Cells[iter, "P"] = vall.oks.value;

                        iter++;
                        #endregion
                    }
                }
            }
            foreach (var vall in corrrectedCads)
            {
                #region Write OKS
                workSheet.Cells[iter, "A"] = vall.cad_num;
                workSheet.Cells[iter, "B"] = vall.oks.cad_num;
                workSheet.Cells[iter, "C"] = vall.oks.type;
                workSheet.Cells[iter, "D"] = vall.oks.name;
                workSheet.Cells[iter, "E"] = vall.oks.adress;
                workSheet.Cells[iter, "F"] = vall.oks.ownership;
                workSheet.Cells[iter, "G"] = vall.oks.summaryArea;
                workSheet.Cells[iter, "H"] = vall.oks.minFloors;
                workSheet.Cells[iter, "I"] = vall.oks.maxFloors;
                workSheet.Cells[iter, "J"] = vall.oks.numsOfUndergroundFloor;
                workSheet.Cells[iter, "K"] = vall.oks.function;
                workSheet.Cells[iter, "L"] = vall.oks.years;
                workSheet.Cells[iter, "P"] = vall.oks.value;

                iter++;
                #endregion
            }
            iter = 2;
            foreach (var val in reSearch)
            {
                workSheet.Cells[iter, "O"] = val;
                iter++;
            }
        }

        static void parser(List<Elem> elem, string driverPath)
        {

            int cnt = 0;

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

            for (int i = 0; i < elem.Count; i++)
            {
                string cad_num = elem[i].cad_num;
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

                        while (cntr < 20 && !equal)
                        {
                            driver.FindElement(By.CssSelector(@"#app-search-submit")).Click();
                            Thread.Sleep(2500);
                            pane = driver.FindElements(By.CssSelector(@"#feature-oks-info > div"));
                            cad_numFromPane = Regex.Match(val, @"Кад. номер:#([^#]+)#", RegexOptions.Compiled).Groups[1].Value;
                            equal = cad_num.Equals(cad_numFromPane);
                            cntr++;
                            if (equal) break;
                            Console.WriteLine("тупняк " + cntr);
                        }
                        cntr = 0;


                        if (!condition.Equals("-") && equal)
                        {
                            OKS oks = new OKS(val, cad_num);
                            elem[i].oks = oks;

                            break;
                        }
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
                    var name = e.GetType().Name;
                    if (name.Equals("ArgumentOutOfRangeException") ||
                        name.Equals("WebDriverTimeoutException"))
                    {
                        OKS oks = new OKS(cad_num, e.GetType().Name, -999);
                        elem[i].oks = oks;
                    }
                    else
                    {
                        OKS oks = new OKS(cad_num, e.GetType().Name, -999);
                        elem[i].oks = oks;
                    }
                    //OKS oks = new OKS(cad_num, e.GetType().Name, "", "", "", "", "", "", "", "");
                    //elem[i].oks = oks;
                }
            }
            driver.Close();
            //return elem;
        }

        static List<Elem> reSearchParser(List<string> reSearch, string driverPath)
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
                            driver.FindElement(By.CssSelector(@"#app-search-submit")).Click();
                            pane = driver.FindElements(By.CssSelector(@"#feature-oks-info > div"));
                            cad_numFromPane = Regex.Match(val, @"Кад. номер:#([^#]+)#", RegexOptions.Compiled).Groups[1].Value;
                            equal = cad_num.Equals(cad_numFromPane);
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
                            OKS oks = new OKS(cad_num, e.GetType().Name, -999);
                            result.Add(new Elem(reSearch[i], oks));
                        }
                        else
                        {
                            OKS oks = new OKS(cad_num, e.GetType().Name, -999);
                            result.Add(new Elem(reSearch[i], oks));
                        }
                    }
                }
            }
            driver.Close();
            return result;
        }

        static void parserTest(string cad_num, string driverPath)
        {

            int cnt = 0;

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

                        break;
                    }
                    else if (counter == 19)
                    {
                        OKS oks = new OKS(val, cad_num);

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
                var name = e.GetType().Name;
                if (name.Equals("ArgumentOutOfRangeException") ||
                    name.Equals("WebDriverTimeoutException"))
                {
                    OKS oks = new OKS(cad_num, e.GetType().Name, -999);

                }
                else
                {
                    OKS oks = new OKS(cad_num, e.GetType().Name, -999);

                }
                //OKS oks = new OKS(cad_num, e.GetType().Name, "", "", "", "", "", "", "", "");
                //elem[i].oks = oks;
            }

            driver.Close();
            //return elem;
        }

        static void ExcelFromValue()
        {
            string path = @"C:\Users\vtsvetkov\source\repos\pkk_5_parser\pkk_5_parser\resourses\Correct.xlsx";
            Application app = new Application();
            app.Visible = true;
            app.Workbooks.Open(path);
            Worksheet workSheet = app.ActiveSheet;

            int iter = 2;

            workSheet.Cells[1, "A"] = "CAD_NUM_Original";
            workSheet.Cells[1, "B"] = "CAD_NUM";
            workSheet.Cells[1, "C"] = "Тип";
            workSheet.Cells[1, "D"] = "Наименование";
            workSheet.Cells[1, "E"] = "Адрес";
            workSheet.Cells[1, "F"] = "Форма собственности";
            workSheet.Cells[1, "G"] = "Общая площадь";
            workSheet.Cells[1, "H"] = "Общая этажность";
            workSheet.Cells[1, "I"] = "Подземная этажность";
            workSheet.Cells[1, "J"] = "Назначение";
            workSheet.Cells[1, "L"] = "IncorrectCads";
            workSheet.Cells[1, "P"] = "Value";

            workSheet.Range["A1", "J1"].Font.Bold = 1;
            workSheet.Columns["A:P"].AutoFit();


            while (workSheet.Cells[iter, "A"].Value != null)
            {
                var val = workSheet.Cells[iter, "P"].Value;
                var cad = workSheet.Cells[iter, "A"].Value;
                if (val != null)
                {
                    OKS oks = new OKS(val, cad);

                    workSheet.Cells[iter, "A"] = cad;
                    workSheet.Cells[iter, "B"] = oks.cad_num;
                    workSheet.Cells[iter, "C"] = oks.type;
                    workSheet.Cells[iter, "D"] = oks.name;
                    workSheet.Cells[iter, "E"] = oks.adress;
                    workSheet.Cells[iter, "F"] = oks.ownership;
                    workSheet.Cells[iter, "G"] = oks.summaryArea;
                    workSheet.Cells[iter, "H"] = oks.minFloors;
                    workSheet.Cells[iter, "I"] = oks.maxFloors;
                    workSheet.Cells[iter, "J"] = oks.numsOfUndergroundFloor;
                    workSheet.Cells[iter, "K"] = oks.function;
                    workSheet.Cells[iter, "L"] = oks.years;
                    workSheet.Cells[iter, "P"] = oks.value;
                }
                Console.WriteLine(iter);
                iter++;
            }
        }
    }






}