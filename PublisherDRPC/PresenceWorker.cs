using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;

using PublisherDRPC.Core;

using MBDRPC.Helpers;

namespace PublisherDRPC
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
            Timer = new Timer(_ => CheckMicrosoftPublisher(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
        }


        private void CheckMicrosoftPublisher()
        {
            processName = "MSPUB"; // Default value

            var isRunning = RunningAppChecker.IsAppRunning(processName);

            if (isRunning)
            {
                if (isFirstRun)
                {
                    officeAppSubscriptionType = GetOfficeVersion();

                    presence.InitializePresence("1224081191947210772");

                    presence.UpdateLargeImage("publisherlogo", "Microsoft Publisher");
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
        /// Checks if any Microsoft Publisher composition/Window is open
        /// </summary>
        private static bool IsAnyOpenWindow()
        {
            // Check if Microsoft Publisher is running
            var processes = Process.GetProcessesByName( "MSPUB" )
                                   .Where( p => ! string.IsNullOrEmpty( p.MainWindowTitle ) );

            return processes.Any();
        }


        /// <summary>
        /// Gets the names of all open Compositions/Windows in Microsoft Publisher
        /// </summary>
        private static string[] GetPublisherOpenWindowNames()
        {
            // Retrieve the names of all open compositions/windows in Microsoft Publisher
            var windowNames = Process.GetProcessesByName( "MSPUB" )
                                     .Where( p => ! string.IsNullOrEmpty( p.MainWindowTitle ) )
                                     .Select( p => p.MainWindowTitle.Replace( " - Publisher" , "" ) )
                                     .ToArray();

            return windowNames;
        }


        /// <summary>
        /// Checks if the Microsoft Publisher home screen window is active
        /// </summary>
        private static bool IsHomeScreenActive()
        {
            var openWindowNames = GetPublisherOpenWindowNames();

            if (openWindowNames.Length <= 0) return false;

            var windowName = openWindowNames[0];
            return ! ( windowName.EndsWith( " - Publisher" ) );
        }

        /// <summary>
        /// Checks if the Microsoft Publisher home screen window is active
        /// </summary>
        private static bool IsHomeScreenActive(IReadOnlyList<string> openWindowNames)
        {
            if (openWindowNames.Count <= 0) return false;

            var windowName = openWindowNames[0];
            return windowName.Equals( "Publisher" );
        }




        /// <summary>
        /// Updates the presence
        /// </summary>
        private void UpdatePresence()
        {
            //Check if any composition is open
            if (IsAnyOpenWindow())
            {
                var openWindowNames = GetPublisherOpenWindowNames();

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