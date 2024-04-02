using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

using MBDRPC.Core;
using MBDRPC.Helpers;

namespace AccessDRPC
{
    public class PresenceWorker
    {
        private Presence presence = new Presence();
        private string officeAppSubscriptionType = "Mirosoft Office";
        private bool isFirstRun = true;
        private DateTime startTime;
        private string processName;

        public Timer Timer;


        /// <summary>
        /// Starts the presence
        /// </summary>
        public void Start()
        {
            Timer = new Timer(_ => CheckMicrosoftAccess(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
        }


        private void CheckMicrosoftAccess()
        {
            processName = "MSACCESS";
            var isPowerPointAppRunning = RunningAppChecker.IsAppRunning(processName);
            if (isPowerPointAppRunning)
            {
                if (isFirstRun)
                {
                    officeAppSubscriptionType = GetOfficeVersion();

                    presence.InitializePresence("1224007046178013196");

                    presence.UpdateLargeImage("accesslogo", "Microsoft Access");
                    presence.UpdateSmallImage("microsoft_365__2022_", GetOfficeVersion());

                    presence.UpdateDetails(officeAppSubscriptionType);


                    startTime = RunningAppChecker.GetProcessStartTime(processName);
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
        /// Checks if any Microsoft Access Database/Window is open
        /// </summary>
        private static bool IsAnyOpenWindow()
        {
            // Check if Microsoft Access is running
            var processes = Process.GetProcessesByName("MSACCESS")
                                   .Where(p => !string.IsNullOrEmpty(p.MainWindowTitle) &&
                                                p.MainWindowTitle != "Microsoft Access" &&
                                                p.MainWindowTitle != "Access");

            return processes.Any();
        }


        /// <summary>
        /// Gets the names of all open Databases/Windows in Microsoft Access
        /// </summary>
        private static string[] GetAccessOpenWindowNames()
        {
            // Retrieve the names of all open databases/windows in Microsoft Access
            var windowsNames = new ConcurrentBag<string>();

            var processes = Process.GetProcessesByName("MSACCESS")
                                   .Where(p => !string.IsNullOrEmpty(p.MainWindowTitle) );

            Parallel.ForEach(processes, process =>
            {
                // Access the process main window title and get only the part between Access - and : path
                var mainWindowTitle = process.MainWindowTitle;

                if (mainWindowTitle.Contains("-"))
                {
                    var startIndex = process.MainWindowTitle.IndexOf(" - " , StringComparison.Ordinal ) + 2;
                    var endIndex   = process.MainWindowTitle.IndexOf(":" , StringComparison.Ordinal );
                    var length     = endIndex - startIndex - 1;

                    mainWindowTitle = process.MainWindowTitle.Substring(startIndex, length);
                }

                windowsNames.Add(mainWindowTitle);
            });

            return windowsNames.ToArray();
        }


        /// <summary>
        /// Checks if the Microsoft Access home screen window is active
        /// </summary>
        private static bool IsHomeScreenActive()
        {
            var openWindowNames = GetAccessOpenWindowNames();

            if (openWindowNames.Length <= 0) return false;

            var windowName = openWindowNames[0];
            return !(windowName.EndsWith(" - Microsoft Access") || windowName.EndsWith(" - Access"));
        }






        /// <summary>
        /// Updates the presence
        /// </summary>
        private void UpdatePresence()
        {
            //Check if any database is open
            if (IsAnyOpenWindow())
            {
                var openWindowNames = GetAccessOpenWindowNames();
                var windowName = openWindowNames[0];

                presence.UpdateDetails($"Managing database: {windowName}");
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
            presence.UpdatePresence(); ;
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
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string office365Path = Path.Combine(appDataPath, "Microsoft", "Office");

            string programFilesPath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
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