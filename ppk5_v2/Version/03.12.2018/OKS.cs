using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ppk5_v2
{
    /// <summary>
    /// 
    /// </summary>
    public struct OKS
    {
        public string cad_num;
        public string type;
        public string name;
        public string adress;
        public string ownership;
        public string summaryArea;
        public int minFloors;
        public int maxFloors;
        public string numsOfUndergroundFloor;
        public string years;
        public string function;
        public string value;

        public OKS(string Cad_num, string _type, string Name, string Adress, string Ownership, string SummaryArea,
                    int MinFloors, int MaxFloors, string NumsOfUndergroundFloor, string Year, string Function, string Value)
        {
            cad_num = Cad_num;
            type = _type;
            name = Name;
            adress = Adress;
            ownership = Ownership;
            summaryArea = SummaryArea;
            minFloors = MinFloors;
            maxFloors = MaxFloors;
            numsOfUndergroundFloor = NumsOfUndergroundFloor;
            years = Year;
            function = Function;
            value = Value;
        }

        /**
            Этот конструктор принимает на вход спаршенную строку из росреестра и кадастровый номер.
            Вытаскивает необходимые значения регулярками и заполняет структуру.             
         */
        public OKS(string val, string Cad_num)
        {
            Regex rType = new Regex(@"Тип:#([^#]+)#", RegexOptions.Compiled);
            Regex rName = new Regex(@"Наименование:#([^#]+)#", RegexOptions.Compiled);
            Regex rAdress = new Regex(@"Адрес:#([^#]+)#", RegexOptions.Compiled);
            Regex rOwnership = new Regex(@"Форма собственности:#([-\sа-яА-Я0-9№^&*,.()/\\""'<>@_!?:;]+)#", RegexOptions.Compiled);
            Regex rSummaryArea = new Regex(@"Общая площадь:#([\s\d]+[,]?[\s\d]*)*.+#", RegexOptions.Compiled);
            Regex rNumsOfFloors = new Regex(@"общая этажность:#([-\s0-9]+)#", RegexOptions.Compiled);
            Regex rNumsOfUndergroundFloor = new Regex(@"подземная этажность:#([-\s0-9]+)#", RegexOptions.Compiled);
            Regex rFunction = new Regex(@"Назначение:#([^#]+)#", RegexOptions.Compiled);
            Regex rYear = new Regex(@"ввод в эксплуатацию:#([^#]+)#", RegexOptions.Compiled);

            #region Floors
            var numsOfFloors = rNumsOfFloors.Match(val).Groups[1].Value;
            var intFloors = 0;
            //var min = Int32.MaxValue;
            //var max = Int32.MinValue;
            var nums = Regex.Matches(numsOfFloors, @"\d+", RegexOptions.Compiled);
            var tnums = nums.Cast<IEnumerable<string>>();
            
            var min = Int32.Parse(tnums.Min().ToString());
            var max = Int32.Parse(tnums.Max().ToString());
            // TODO: Реализовать через интерфейс IEnumerable Min() Max ()
            foreach (Match fl in nums)
            {
                Console.WriteLine(fl);
                intFloors = Int32.Parse(fl.Value);
                Console.WriteLine(intFloors);
                if (intFloors > max) max = intFloors;
                if (intFloors < min) min = intFloors;
            }
            if (min == Int32.MaxValue) min = 0;
            if (max == Int32.MinValue) max = 0;
            #endregion
            #region UndergroudFloor
            // TODO: После реализоции IEnumerable для общей этажности
            #endregion

            cad_num = Cad_num;
            type = rType.Match(val).Groups[1].Value;
            name = rName.Match(val).Groups[1].Value;
            adress = rAdress.Match(val).Groups[1].Value;
            ownership = rOwnership.Match(val).Groups[1].Value;
            summaryArea = rSummaryArea.Match(val).Groups[1].Value;
            minFloors = min;
            maxFloors = max;
            numsOfUndergroundFloor = rNumsOfUndergroundFloor.Match(val).Groups[1].Value;
            years = rYear.Match(val).Groups[1].Value;
            function = rFunction.Match(val).Groups[1].Value;
            value = val;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="val"></param>
        /// <param name="exceptionName"></param>
        /// <param name="flag"></param>
        public OKS(string val, string exceptionName, int flag)
        {
            cad_num = val;
            type = exceptionName;
            name = null;
            adress = null;
            ownership = null;
            summaryArea = null;
            minFloors = flag;
            maxFloors = flag;
            numsOfUndergroundFloor = null;
            years = null;
            function = null;
            value = null;
        }


    }
}
