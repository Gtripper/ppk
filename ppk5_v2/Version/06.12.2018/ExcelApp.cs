using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ppk5_v2
{
    class ExcelAppReadData
    {
        private List<Elem> elems;
        private Worksheet worksheet;
        #region Constructors
        public ExcelAppReadData()
        {
        }

        public ExcelAppReadData(string excelPath)
        {
            worksheet = OpenFile(excelPath);
        }

        #endregion
        internal List<Elem> _elems { get => elems; }
        /// <summary>
        /// Открывает файл экселя и формирует шапку
        /// </summary>
        /// <param name="excelPath">Путь к файлу</param>
        /// <returns>Worksheet</returns>
        private Worksheet OpenFile(string excelPath)
        {
            Application app = new Application();
            app.Visible = true;
            app.Workbooks.Open(excelPath);
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

            return workSheet;
        }

        #region ReadData
        /// <summary>
        /// Чтение кадастровых номеров, начиная с ряда 2 и до первой пустой ячейки
        /// </summary>
        /// <param name="columnName">Литера столбца с кадастровыми номерами</param>
        /// <param name="workSheet">Воркшит, из которого производится чтение</param>
        /// <returns>elems</returns>
        private List<Elem> ReadDataFromExcel(string columnName, Worksheet workSheet)
        {
            var iter = 2;
            string cad_num;
            var result = new List<Elem>();
            try
            {
                while (workSheet.Cells[iter, columnName].Value != null)
                {
                    cad_num = (string)(workSheet.Cells[iter, columnName] as Range).Value;
                    if (CheckUpCadNumOKS(cad_num)) result.Add(new Elem(cad_num));
                    iter++;
                }
            }
            catch (Exception e)
            {
                ReadDataFromExcel(columnName, workSheet, iter, result);
                Console.WriteLine("EXCEPTION " + e.GetType().Name + " in row: " + iter);
                Console.WriteLine("Start ReadDataFromExcel at row");
            }
            return result;
        }

        /// <summary>
        /// Чтение кадастровых номеров, начиная с ряда startRow и до первой пустой ячейки
        /// </summary>
        /// <param name="columnName">Литера столбца с кадастровыми номерами</param>
        /// <param name="workSheet">Воркшит, из которого производится чтение</param>
        /// <param name="startRow">Индекс начала поиска</param>
        /// <returns>elems</returns>
        private List<Elem> ReadDataFromExcel(string columnName, Worksheet workSheet, int startRow)
        {
            string cad_num;
            var result = new List<Elem>();
            try
            {
                while (workSheet.Cells[startRow, columnName].Value != null)
                {
                    cad_num = (string)(workSheet.Cells[startRow, columnName] as Range).Value;
                    if (CheckUpCadNumOKS(cad_num)) result.Add(new Elem(cad_num));
                    startRow++;
                }
            }
            catch (Exception e)
            {
                ReadDataFromExcel(columnName, workSheet, startRow, result);
                Console.WriteLine("EXCEPTION " + e.GetType().Name + " in row: " + startRow);
                Console.WriteLine("Start ReadDataFromExcel at row");
            }
            return result;
        }

        /// <summary>
        /// Чтение кадастровых номеров, начиная с ряда startRow и до endRow
        /// </summary>
        /// <param name="columnName">Литера столбца с кадастровыми номерами</param>
        /// <param name="workSheet">Воркшит, из которого производится чтение</param>
        /// <param name="startRow">Индекс начала поиска</param>
        /// <param name="endRow">Индекс окончания поиска</param>
        /// <returns>elems</returns>
        private List<Elem> ReadDataFromExcel(string columnName, Worksheet workSheet, int startRow, int endRow)
        {
            string cad_num;
            var result = new List<Elem>();

            try
            {
                while (startRow < endRow)
                {
                    cad_num = (string)(workSheet.Cells[startRow, columnName] as Range).Value;
                    if (CheckUpCadNumOKS(cad_num)) result.Add(new Elem(cad_num));
                    startRow++;
                }
            }
            catch (Exception e)
            {
                ReadDataFromExcel(columnName, workSheet, startRow, endRow, result);
                Console.WriteLine("EXCEPTION " + e.GetType().Name + " in row: " + startRow);
                Console.WriteLine("Start ReadDataFromExcel at row");
            }
            return result;
        }

        /// <summary>
        /// Чтение кадастровых номеров, начиная с ряда startRow и до первой пустой ячейки (вызывается исключением)
        /// </summary>
        /// <param name="columnName">Литера столбца с кадастровыми номерами</param>
        /// <param name="workSheet">Воркшит, из которого производится чтение</param>
        /// <param name="startRow">Индекс начала поиска</param>
        /// <param name="result">Считанные данные</param>
        /// <returns>elems</returns>
        private List<Elem> ReadDataFromExcel(string columnName, Worksheet workSheet, int startRow, List<Elem> result)
        {
            string cad_num;

            try
            {
                while (workSheet.Cells[startRow, columnName].Value != null)
                {
                    cad_num = (string)(workSheet.Cells[startRow, columnName] as Range).Value;
                    if (CheckUpCadNumOKS(cad_num)) result.Add(new Elem(cad_num));
                    startRow++;
                }
            }
            catch (Exception e)
            {
                ReadDataFromExcel(columnName, workSheet, startRow, result);
                Console.WriteLine("EXCEPTION " + e.GetType().Name + " in row: " + startRow);
                Console.WriteLine("Start ReadDataFromExcel with resoult.count " + result.Count);
            }
            return result;
        }

        /// <summary>
        /// Чтение кадастровых номеров, начиная с ряда startRow и до первой пустой ячейки (вызывается исключением)
        /// </summary>
        /// <param name="columnName">Литера столбца с кадастровыми номерами</param>
        /// <param name="workSheet">Воркшит, из которого производится чтение</param>
        /// <param name="endRow">Индекс окончания поиска</param>
        /// <param name="result">Считанные данные</param>
        /// <returns>elems</returns>
        private List<Elem> ReadDataFromExcel(string columnName, Worksheet workSheet, int startRow, int endRow, List<Elem> result)
        {
            string cad_num;

            try
            {
                while (startRow < endRow)
                {
                    cad_num = (string)(workSheet.Cells[startRow, columnName] as Range).Value;
                    if (CheckUpCadNumOKS(cad_num)) result.Add(new Elem(cad_num));
                    startRow++;
                }
            }
            catch (Exception e)
            {
                ReadDataFromExcel(columnName, workSheet, startRow, endRow, result);
                Console.WriteLine("EXCEPTION " + e.GetType().Name + " in row: " + startRow);
                Console.WriteLine("Start ReadDataFromExcel at row");
            }
            return result;
        }

        /// <summary>
        /// Проверка кадастрового номера на корректность
        /// </summary>
        /// <param name="input">строка, сдержащая кадастровый номер</param>
        /// <returns></returns>
        private bool CheckUpCadNumOKS(string input)
        {
            return Regex.IsMatch(input, @"\d+:\d+:\d+:\d+", RegexOptions.Compiled);
        }
        #endregion

        internal void Run()
        { 
        }

        internal void Run(string columnName)
        {
            elems = ReadDataFromExcel(columnName, worksheet);
        }

        internal void Run(string columnName, int startIndex)
        {
            elems = ReadDataFromExcel(columnName, worksheet, startIndex);
        }

        internal void Run(string columnName, int startIndex, int endIndex)
        {
            elems = ReadDataFromExcel(columnName, worksheet, startIndex, endIndex);
        }

    }

    class ExcelAppWriteData
    {
        private List<Elem> elems;

        internal ExcelAppWriteData() { }

        internal ExcelAppWriteData(IEnumerable<List<Elem>> elem)
        {
            elems = new List<Elem>();
            foreach (var flow in elem)
            {
                foreach (var val in flow)
                {
                    elems.Add(val);
                }
            }
        }

        /// <summary>
        /// Создание пустого файла, формирование шапки
        /// </summary>
        /// <returns>Worksheet</returns>
        private Worksheet CreateFile()
        {
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

            return workSheet;
        }

        private void WriteData(Worksheet worksheet, int startIndex)
        {
            var start = startIndex;
            foreach (var val in elems)
            {
                try
                {
                    #region Write OKS
                    worksheet.Cells[startIndex, "A"] = val.cad_num;
                    worksheet.Cells[startIndex, "B"] = val.oks.cad_num;
                    worksheet.Cells[startIndex, "C"] = val.oks.type;
                    worksheet.Cells[startIndex, "D"] = val.oks.name;
                    worksheet.Cells[startIndex, "E"] = val.oks.adress;
                    worksheet.Cells[startIndex, "F"] = val.oks.ownership;
                    worksheet.Cells[startIndex, "G"] = val.oks.summaryArea;
                    worksheet.Cells[startIndex, "H"] = val.oks.minFloors;
                    worksheet.Cells[startIndex, "I"] = val.oks.maxFloors;
                    worksheet.Cells[startIndex, "J"] = val.oks.numsOfUndergroundFloor;
                    worksheet.Cells[startIndex, "K"] = val.oks.function;
                    worksheet.Cells[startIndex, "L"] = val.oks.years;
                    worksheet.Cells[startIndex, "P"] = val.oks.value;
                }
                catch
                {
                    WriteData(worksheet, start);
                }

                startIndex++;
                #endregion
            }
        }

        internal void Run()
        {
            var wrksht = CreateFile();
            WriteData(wrksht, 2);
        }

        internal void Run(int startIndex)
        {
            var wrksht = CreateFile();
            WriteData(wrksht, startIndex);
        }
    }
}
