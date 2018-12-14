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
        public float summaryArea;
        public int minFloors;
        public int maxFloors;
        public int numsOfUndergroundFloor;
        public string years;
        public string function;
        public string value;

        public OKS(string Cad_num, string _type, string Name, string Adress, string Ownership, float SummaryArea,
                    int MinFloors, int MaxFloors, int NumsOfUndergroundFloor, string Year, string Function, string Value)
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
            var nums = Regex.Matches(numsOfFloors, @"\d+", RegexOptions.Compiled);
            var list = nums.Cast<Match>().Select(match => Int32.Parse(match.Value)).ToList();
            if (list.Count > 0)
            {
                minFloors = list.Min();
                maxFloors = list.Max();
            }
            else
            {
                minFloors = 0;
                maxFloors = 0;
            }
            #endregion
            #region UndergroudFloor
            var numsOfUnregroudFloors = rNumsOfUndergroundFloor.Match(val).Groups[1].Value;
            nums = Regex.Matches(numsOfUnregroudFloors, @"\d+", RegexOptions.Compiled);
            list = nums.Cast<Match>().Select(match => Int32.Parse(match.Value)).ToList();
            if (list.Count > 0)
            {
                numsOfUndergroundFloor = list.Max();
            }
            else
            {
                numsOfUndergroundFloor = 0;
            }
            #endregion
            #region summaryArea
            try
            {
                summaryArea = Single.Parse(rSummaryArea.Match(val).Groups[1].Value);
            }
            catch
            {
                summaryArea = 0;
            }
            #endregion
            cad_num = Cad_num;
            type = rType.Match(val).Groups[1].Value;
            name = rName.Match(val).Groups[1].Value;
            adress = rAdress.Match(val).Groups[1].Value;
            ownership = rOwnership.Match(val).Groups[1].Value;            
            years = rYear.Match(val).Groups[1].Value;
            function = rFunction.Match(val).Groups[1].Value;
            value = val;
        }

        public OKS(string val, string exceptionName, int flag)
        {
            cad_num = val;
            type = exceptionName;
            name = null;
            adress = null;
            ownership = null;
            summaryArea = 0;
            minFloors = flag;
            maxFloors = flag;
            numsOfUndergroundFloor = 0;
            years = null;
            function = null;
            value = null;
        }


    }
}
