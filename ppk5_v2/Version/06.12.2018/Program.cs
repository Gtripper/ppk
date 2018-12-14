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
            var excelPath = @"C:\Users\vtsvetkov\source\repos\pkk_5_parser\pkk_5_parser\resourses\Test.xlsx";
            var driverPath = @"C:\Users\vtsvetkov\source\repos\pkk_5_parser";
            var numOfThreads = 10;
            var threadLenght = 5;

            IFabric fab = new Fabric(excelPath, driverPath, numOfThreads, threadLenght);
            fab.SearchOKS("A", 2);
        }
    }
}