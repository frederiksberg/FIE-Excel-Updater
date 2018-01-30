using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelUpdater
{
    class Program
    {
        private static Application _excel;

        private static System.Timers.Timer _timer;

        private static Queue<string> _fileLocations = new Queue<string>();
        private static DateTime Current_EndDate;
        private static Workbook Current_WorkBook;

        private static event EventHandler finishedWorkBook;

        static void Main(string[] args)
        {
            //Start Excel and disable alerts
            _excel = new Application();
            _excel.DisplayAlerts = false;
            _excel.Visible = true;

            //Read in all files
            _fileLocations = new Queue<string>(Directory.GetFiles("C:\\Users\\nibr02\\Documents\\excel-test"));


            finishedWorkBook += Program_finishedWorkBook;
            UpdateWorkBook();

            //Stop program from ending
            Console.ReadLine();
        }

        private static void Program_finishedWorkBook(object sender, EventArgs e)
        {
            if (_fileLocations.Count > 0)
            {
                UpdateWorkBook();
            }
            else
            {
                _excel.DisplayAlerts = true;
                _excel.Quit();              
                Marshal.ReleaseComObject(_excel);
                Environment.Exit(0);
            }
        }

        private static void UpdateWorkBook()
        {
            var file = _fileLocations.Dequeue();
            Current_WorkBook = _excel.Workbooks.Open(file);

            //Find the current end date from the workbook
            Current_EndDate = DateTime.Parse(Current_WorkBook.Names.Item("GetDataEnd").RefersToRange.Value);

            Current_WorkBook.SheetChange += SheetCalculate;

            //Give focus to the current workbook
            Current_WorkBook.Activate();

            //Use send keys to active prisme update through shortcuts
            _excel.SendKeys("%");
            _excel.SendKeys("Ø");
            _excel.SendKeys("y3");
            _excel.SendKeys("o");
        }

        private static void SheetCalculate(object sender, Range Target)
        {
            //See if process end date has been changed (which means we're done calculating)
            var newDate = DateTime.Parse(Current_WorkBook.Names.Item("GetDataEnd").RefersToRange.Value);
            if (newDate != Current_EndDate)
            {
                Current_WorkBook.SheetChange -= SheetCalculate;
                Current_WorkBook.AfterSave += Current_WorkBook_AfterSave;
                Current_WorkBook.Save();                              
            }
        }

        private static void Current_WorkBook_AfterSave(bool Success)
        {
            Current_WorkBook.RefreshAll();

            //We have to wait for the prisme add-in bacause it's doing its job on another thread
            _timer = new System.Timers.Timer();
            _timer.Elapsed += _timer_Elapsed; 
            _timer.Interval = 1000;
            _timer.Enabled = true;           
        }

        private static void _timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            _timer.Enabled = false;
            _timer.Elapsed -= _timer_Elapsed;
            finishedWorkBook(null, new EventArgs());
        }
    }
}
