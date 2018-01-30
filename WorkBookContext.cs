using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUpdater
{
    class WorkBookContext
    {
        public Workbook WorkBook { get; set; }

        public DateTime StartTime { get; set; }
    }
}
