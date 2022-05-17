using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelData
{
    /// <summary>
    /// Глобальный список с пропусками
    /// </summary>
    public class YearlySkips
    {
        //студент месяц и все виды пропусков
        public static List<(string student, string month, int allSkips, int okSkips, int notOkSkips)> Data = new List<(string student, string month, int allSkips, int okSkips, int notOkSkips)>();
    }
}
