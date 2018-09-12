using Microsoft.Office.Interop.Excel;
using OutLook = Microsoft.Office.Interop.Outlook;
using NLog;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace ExcelUpdater
{
    class Program
    {
        private static Application _excel;

        private static System.Timers.Timer _timer;
        private static Logger _logger = LogManager.GetCurrentClassLogger();

        private static Queue<string> _fileLocations = new Queue<string>();
        private static DateTime Current_EndDate;
        private static Workbook Current_WorkBook;

        private static event EventHandler finishedWorkBook;

        static void Main(string[] args)
        {
            Console.WindowWidth = 110;


            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

            _logger.Info("Starting Excel");

            //Start Excel and disable alerts
            _excel = new Application();
            _excel.DisplayAlerts = false;
            _excel.Visible = true;
            _excel.WindowState = XlWindowState.xlNormal;

            _logger.Info("Excel started successfully");

            var path = ConfigurationManager.AppSettings["BaseFolderLocation"];

            ReadFiles(path);

            finishedWorkBook += Program_finishedWorkBook;
            UpdateWorkBook();

            //Stop program from ending
            Console.ReadLine();

            SystemEvents.SessionEnded += SystemEvents_SessionEnded;
            SystemEvents.SessionEnding += SystemEvents_SessionEnding; 
            SystemEvents.SessionSwitch += SystemEvents_SessionSwitch;

        }

        private static void SystemEvents_SessionEnding(object sender, SessionEndingEventArgs e)
        {
            _logger.Info($"Session ended because {e.Reason} ");
        }

        private static void SystemEvents_SessionSwitch(object sender, SessionSwitchEventArgs e)
        {
            _logger.Info($"Session switched because {e.Reason} ");
        }

        private static void SystemEvents_SessionEnded(object sender, SessionEndedEventArgs e)
        {
            _logger.Info($"Session ended because {e.Reason} ");
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            _logger.Error(e.ExceptionObject.ToString());
        }

        private static void ReadFiles(string path)
        {
            _logger.Info($"Finding all Excel workbooks in {path}");

            if (!Directory.Exists(path))
            {
                _logger.Error($"{path} does not exist");
                Environment.Exit(1);
            }

            //Read in all files
            var files = new DirectoryInfo(path).GetFiles().Where(x => !x.Attributes.HasFlag(FileAttributes.Hidden));


            _fileLocations = new Queue<string>(files.Select(x => x.FullName));

            foreach (var file in _fileLocations.ToList())
            {
                _logger.Info($"{file} added for processing");
            }
        }

        private static void Program_finishedWorkBook(object sender, EventArgs e)
        {
            if (_fileLocations.Count > 0)
            {
                UpdateWorkBook();
            }
            else
            {
                _logger.Info($"Finished processing all workbooks");
                _logger.Info($"Shutting down Excel");
                _excel.Quit();
                _excel.DisplayAlerts = true;
                Marshal.ReleaseComObject(_excel);

                //SendEmailLog();

                Environment.Exit(0);
            }
        }

        private static void SendEmailLog()
        {
            OutLook.Application app = new OutLook.Application();
            OutLook.MailItem mailItem = app.CreateItem(OutLook.OlItemType.olMailItem);
            mailItem.Subject = "Excel Updater Log";
            mailItem.To = "nibr02@frederiksberg.dk";
            mailItem.Body = "This is the message.";
            mailItem.Attachments.Add(AppDomain.CurrentDomain.BaseDirectory + "log.txt");//logPath is a string holding path to the log.txt file
            mailItem.Importance = OutLook.OlImportance.olImportanceHigh;
            mailItem.Display(false);
        }

        private static void UpdateWorkBook()
        {
            var file = _fileLocations.Dequeue();

            _logger.Info($"{file} began processing");

            try
            {
                Current_WorkBook = _excel.Workbooks.Open(file, UpdateLinks:2);

                _logger.Info($"{Current_WorkBook.Name} opened in Excel");

                //Find the current end date from the workbook
                Current_EndDate = DateTime.Parse(Current_WorkBook.Names.Item("GetDataEnd").RefersToRange.Value.ToString());

                Current_WorkBook.SheetChange += SheetCalculate;

                //_excel.ActiveWorkbook.Activate();
                //_excel.ActiveWorkbook.ActiveSheet.Select();

                User32.SetForegroundWindow((IntPtr)_excel.Hwnd);

                _logger.Info($"{Current_WorkBook.Name} updated started");



                //Use send keys to active prisme update through shortcuts
                _excel.SendKeys("%");
                _excel.SendKeys("X");
                _excel.SendKeys("Y");
                _excel.SendKeys("O");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"An error occured while running update on the workbook: {Current_WorkBook.Name}");
            }

        }

        private static void SheetCalculate(object sender, Range Target)
        {
            try
            {
                //See if process end date has been changed (which means we're done calculating)
                var newDate = DateTime.Parse(Current_WorkBook.Names.Item("GetDataEnd").RefersToRange.Value.ToString());
                if (newDate != Current_EndDate)
                {
                    _logger.Info($"{Current_WorkBook.Name} updated finished");
                    Current_WorkBook.SheetChange -= SheetCalculate;
                    Current_WorkBook.AfterSave += Current_WorkBook_AfterSave;
                    Current_WorkBook.Save();
                }
            }         
            catch(Exception ex)
            {
                _logger.Error(ex, $"An error occured while checking update status on the workbook: {Current_WorkBook.Name}");
            }
        }

        private static void Current_WorkBook_AfterSave(bool Success)
        {
            _logger.Info($"{Current_WorkBook.Name} saved successfully");

            Current_WorkBook.RefreshAll();

            //We have to wait for the prisme add-in bacause it's doing its job on another thread
            _timer = new System.Timers.Timer();
            _timer.Elapsed += _timer_Elapsed;
            _timer.Interval = 1000;
            _timer.Enabled = true;
        }

        private static void _timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            _logger.Info($"{Current_WorkBook.Name} processing complete");
            _timer.Enabled = false;
            _timer.Elapsed -= _timer_Elapsed;
            Current_WorkBook.Close();
            finishedWorkBook(null, new EventArgs());
        }
    }
}
