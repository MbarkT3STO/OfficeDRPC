using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using ExcelDRPC.Core;
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
        /// Checks if any Microsoft Excel workbook/window is open
        /// </summary>
        private static bool IsAnyOpenWindow()
        {
            // Check if Microsoft Excel workbook is open
            var processes = Process.GetProcessesByName( "EXCEL" )
                                   .Where( p => ! string.IsNullOrEmpty( p.MainWindowTitle ) );

            return processes.Any();
        }


        /// <summary>
        /// Gets the names of all open workbooks/windows in Microsoft Excel
        /// </summary>
        private static string[] GetExcelOpenWindowNames()
        {
            // Retrieve the names of all open workbooks/windows in Microsoft Excel
            var windowNames = Process.GetProcessesByName( "EXCEL" )
                                     .Where( p => ! string.IsNullOrEmpty( p.MainWindowTitle ) )
                                     .Select( p => p.MainWindowTitle.Replace( " - Excel" , "" ) )
                                     .ToArray();

            return windowNames;
        }


        /// <summary>
        /// Checks if the Microsoft Excel home screen window is active
        /// </summary>
        private static bool IsHomeScreenActive()
        {
            var openWindowNames = GetExcelOpenWindowNames();

            if (openWindowNames.Length <= 0) return false;

            var windowName = openWindowNames[0];
            return !(windowName.EndsWith(" - Excel"));
        }

        /// <summary>
        /// Checks if the Microsoft Excel home screen window is active
        /// </summary>
        private static bool IsHomeScreenActive(IReadOnlyList<string> openWindowNames)
        {
            if (openWindowNames.Count <= 0) return false;

            var windowName = openWindowNames[0];
            return windowName.Equals( "Excel" );
        }




        /// <summary>
        /// Updates the presence
        /// </summary>
        private void UpdatePresence()
        {
            //Check if any workbook is open
            if (IsAnyOpenWindow())
            {
                var openWindowNames = GetExcelOpenWindowNames();

                if (IsHomeScreenActive(openWindowNames))
                {
                    presence.UpdateDetails("Home screen");
                }
                else
                {
                    if (openWindowNames.Length > 0)
                    {
                        var windowName = openWindowNames[0];

                        presence.UpdateDetails($"Editing: {windowName}");
                    }
                }
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