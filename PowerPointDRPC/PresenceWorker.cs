using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using MBDRPC.Core;
using MBDRPC.Helpers;

namespace PowerPointDRPC
{
	public class PresenceWorker
    {
        private Presence presence                  = new Presence();
        private string   officeAppSubscriptionType = "Mirosoft Office";
        private bool     isFirstRun                = true;
        private DateTime startTime;
        private string   processName;

        public Timer Timer;


        /// <summary>
        /// Starts the presence
        /// </summary>
        public void Start()
		{
            Timer = new Timer(_ => CheckMicrosoftPowerPoint(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
        }


        private void CheckMicrosoftPowerPoint()
        {
            processName = "POWERPNT";
            var isPowerPointAppRunning = RunningAppChecker.IsAppRunning(processName);
            if (isPowerPointAppRunning)
            {
                if (isFirstRun)
                {
                    officeAppSubscriptionType = GetOfficeVersion();

                    presence.InitializePresence("1224001855395463208");

                    presence.UpdateLargeImage("powerpointlogo", "Microsoft PowerPoint");
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
        /// Checks if any Microsoft PowerPoint presentation/Window is open
        /// </summary>
        private static bool IsAnyOpenWindow()
        {
            // Check if Microsoft PowerPoint is running
            var processes = Process.GetProcessesByName( "POWERPNT" )
                                   .Where( p => ! string.IsNullOrEmpty( p.MainWindowTitle ) );

            return processes.Any();
        }


        /// <summary>
        /// Gets the names of all open Presentations/Windows in Microsoft PowerPoint
        /// </summary>
        private static string[] GetPowerPointOpenWindowNames()
        {
            // Retrieve the names of all open presentations/windows in Microsoft PowerPoint
            var windowNames = Process.GetProcessesByName("POWERPNT")
                                     .Where(p => !string.IsNullOrEmpty(p.MainWindowTitle))
                                     .Select(p => p.MainWindowTitle.Replace(" - PowerPoint", ""))
                                     .ToArray();

            return windowNames;
        }


        /// <summary>
        /// Checks if the Microsoft PowerPoint home screen window is active
        /// </summary>
        private static bool IsHomeScreenActive()
        {
            var openWindowNames = GetPowerPointOpenWindowNames();

            if (openWindowNames.Length <= 0) return false;

            var windowName = openWindowNames[0];
            return ! ( windowName.EndsWith( " - PowerPoint" ) );
        }


        /// <summary>
        /// Checks if the Microsoft PowerPoint home screen window is active
        /// </summary>
        private static bool IsHomeScreenActive(IReadOnlyList<string> openWindowNames)
        {
            if (openWindowNames.Count <= 0) return false;

            var windowName = openWindowNames[0];
            return windowName.Equals("PowerPoint");
        }





        /// <summary>
        /// Updates the presence
        /// </summary>
        private void UpdatePresence()
        {
            //Check if any presentation is open
            if (IsAnyOpenWindow())
            {
                var openWindowNames = GetPowerPointOpenWindowNames();

                if (IsHomeScreenActive(openWindowNames))
                {
                    presence.UpdateDetails("Home screen");
                }
                else
                {
                    var windowName = openWindowNames[0];

                    presence.UpdateDetails($"Editing: {windowName}");
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