using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using MBDRPC.Core;
using MBDRPC.Helpers;

namespace ExcelDRPC
{
	public class PresenceWorker
    {
        private string   officeAppSubscriptionType = "Mirosoft Office";
		private Presence presence                  = new Presence();
        private bool     isFirstRun                = true;
        private DateTime startTime;
        private string   processName;
        public  Timer    Timer;


        /// <summary>
        /// Starts the presence
        /// </summary>
        public void Start()
		{
            Timer = new Timer(_ => CheckMicrosoftExcel(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
        }


        private void CheckMicrosoftExcel()
        {
            var isMsExcelRunning = RunningAppChecker.IsAppRunning("excel");
            if (isMsExcelRunning)
            {
                if (isFirstRun)
                {
                    processName    = "EXCEL";
                    officeAppSubscriptionType = GetOfficeVersion();

                    presence.InitializePresence("1223982816459489350");

                    presence.UpdateLargeImage("exccellogo", "Microsoft Excel");
                    presence.UpdateSmallImage("microsoft_365__2022_", officeAppSubscriptionType);

                    presence.UpdateDetails(officeAppSubscriptionType);


                    startTime  = RunningAppChecker.GetProcessStartTime(processName);
                    isFirstRun = false;
                }

                UpdatePresence();
            }
            else
            {
                presence.ShutDown();
                isFirstRun = true;
            }
        }






        /// <summary>
        /// Checks if any Microsoft Excel workbook is open
        /// </summary>
        private static bool IsAnyOpenWorkbook()
        {
            // Check if Microsoft Excel workbook is open
            var processes = Process.GetProcessesByName( "EXCEL" )
                                   .Where( p => ! string.IsNullOrEmpty( p.MainWindowTitle ) &&
                                                p.MainWindowTitle != "Excel"                &&
                                                p.MainWindowTitle != "Microsoft Excel" );

            return processes.Any();
        }


        /// <summary>
        /// Gets the names of all open workbooks in Microsoft Excel
        /// </summary>
        private static string[] GetExcelOpenWorkbookNames()
        {
            // Retrieve the names of all open workbooks in Microsoft Excel
            var workbookNames = new ConcurrentBag<string>();

            var processes = Process.GetProcessesByName( "EXCEL" )
                                   .Where( p => !string.IsNullOrEmpty( p.MainWindowTitle ));

            Parallel.ForEach( processes , process =>
                                          {
                                              // Access the process main window title and remove the " - Microsoft Excel" or " - Excel" suffix
                                              var mainWindowTitle = process.MainWindowTitle.Replace(" - Microsoft Excel", "").Replace(" - Excel", "");

                                              workbookNames.Add( mainWindowTitle );
                                          } );

            return workbookNames.ToArray();
        }


        /// <summary>
        /// Checks if the Microsoft Excel home screen is active
        /// </summary>
        private static bool IsHomeScreenActive()
        {
            var workbookNames = GetExcelOpenWorkbookNames();

            if (workbookNames.Length <= 0) return false;

            var currentWorkbookName = workbookNames[0];
            return !(currentWorkbookName.EndsWith(" - Microsoft Excel") || currentWorkbookName.EndsWith(" - Excel"));
        }






        /// <summary>
        /// Updates the presence
        /// </summary>
        private void UpdatePresence()
        {
            //Check if any workbook is open
            if (IsAnyOpenWorkbook())
            {
                var workbookNames       = GetExcelOpenWorkbookNames();
                var currentWorkbookName = workbookNames[0];

                presence.UpdateDetails($"Editing workbook: {currentWorkbookName}");
            }
            else if (IsHomeScreenActive())
            {
                presence.UpdateDetails("Home screen");
            }
            else
            {
                presence.UpdateDetails(officeAppSubscriptionType);
            }

            UpdatePresenceTime();
            presence.UpdatePresence();
        }


        /// <summary>
        /// Updates the presence time
        /// </summary>
        private void UpdatePresenceTime()
        {
            var elapsedTime = (DateTime.Now - startTime).ToString(@"hh\:mm\:ss");
            presence.UpdateState(elapsedTime);
        }


        public static string GetOfficeVersion()
        {
            string appDataPath   = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string office365Path = Path.Combine(appDataPath, "Microsoft", "Office");

            string programFilesPath    = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            string perpetualOfficePath = Path.Combine(programFilesPath, "Microsoft Office", "root", "Office16");

            if (Directory.Exists(office365Path))
            {
                return "Microsoft 365";
            }
            else if (Directory.Exists(perpetualOfficePath))
            {
                return "Microsoft Office";
            }
            else
            {
                return "Microsoft Office";
            }
        }

    }
}