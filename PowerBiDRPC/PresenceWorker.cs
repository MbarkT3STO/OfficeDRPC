using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

using MBDRPC.Core;
using MBDRPC.Helpers;

namespace PowerBiDRPC
{
    public class PresenceWorker
    {
        private string officeAppSubscriptionType = "Mirosoft Office";
        private Presence presence = new Presence();
        private bool isFirstRun = true;
        private DateTime startTime;
        private string processName;
        public Timer Timer;


        /// <summary>
        /// Starts the presence
        /// </summary>
        public void Start()
        {
            Timer = new Timer(_ => CheckMicrosoftPowerBI(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
        }


        private void CheckMicrosoftPowerBI()
        {
            var isPowerBIRunning = RunningAppChecker.IsAppRunning("PBIDesktop");
            if (isPowerBIRunning)
            {
                if (isFirstRun)
                {
                    processName = "PBIDesktop";
                    officeAppSubscriptionType = GetOfficeVersion();

                    presence.InitializePresence("1224821008234709153");

                    presence.UpdateLargeImage("powerbilogo", "Microsoft Power BI");
                    presence.UpdateSmallImage("officelogo2", officeAppSubscriptionType);

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
        /// Checks if any Microsoft Power BI Report/window is open
        /// </summary>
        private static bool IsAnyOpenWindow()
        {
            // Check if Microsoft Power BI Report is open
            var processes = Process.GetProcessesByName("PBIDesktop")
                                   .Where(p => !string.IsNullOrEmpty(p.MainWindowTitle));

            return processes.Any();
        }


        /// <summary>
        /// Gets the names of all open reports/windows in Microsoft Power BI
        /// </summary>
        private static string[] GetPowerBIOpenWindowNames()
        {
            // Retrieve the names of all open reports/windows in Microsoft Power BI
            var windowNames = Process.GetProcessesByName("PBIDesktop")
                                     .Where(p => !string.IsNullOrEmpty(p.MainWindowTitle))
                                     .Select(p => p.MainWindowTitle)
                                     .ToArray();

            return windowNames;
        }


        /// <summary>
        /// Checks if the Microsoft Power BI home screen window is active
        /// </summary>
        private static bool IsHomeScreenActive()
        {
            var openWindowNames = GetPowerBIOpenWindowNames();

            if (openWindowNames.Length <= 0) return false;

            var windowName = openWindowNames[0];
            return (windowName.Equals("Untitled - Power BI Desktop") || windowName.Equals("Power BI Desktop"));
        }


        /// <summary>
        /// Checks if the Microsoft Power BI home screen window is active
        /// </summary>
        private static bool IsHomeScreenActive(IReadOnlyList<string> openWindowNames)
        {
            if (openWindowNames.Count <= 0) return false;

            var windowName = openWindowNames[0];
            return (windowName.Equals("Untitled - Power BI Desktop") || windowName.Equals("Power BI Desktop"));
        }


        /// <summary>
        /// Updates the presence
        /// </summary>
        private void UpdatePresence()
        {
            //Check if the home screen is active
            if (IsAnyOpenWindow())
            {
                var openWindowNames = GetPowerBIOpenWindowNames();

                if (IsHomeScreenActive(openWindowNames))
                {
                    presence.UpdateDetails("Home screen");
                }
                else
                {
                    var windowName = openWindowNames
                       .FirstOrDefault( s => ! s.Equals( "Untitled - Power BI Desktop" ) &&
                                             ! s.Equals( "Power BI Desktop" ) );

                   presence.UpdateDetails( $"Editing: {windowName}" );
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