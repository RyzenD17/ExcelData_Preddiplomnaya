using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelData
{
    public class StudentsData
    {
        public string FIO { get;set; }
        //перое это дата а второе должно было быть часами пропуска но не вышло
        public List<(DateTime, int?)> data { get; set; }
        
    }
}
