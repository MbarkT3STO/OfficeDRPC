using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

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
        /// Checks if any Microsoft PowerPoint presentation is open
        /// </summary>
        private static bool IsAnyOpenPresentation()
        {
            // Check if Microsoft PowerPoint is running
            var processes = Process.GetProcessesByName( "POWERPNT" )
                                   .Where( p => ! string.IsNullOrEmpty( p.MainWindowTitle ) &&
                                                p.MainWindowTitle != "PowerPoint"           &&
                                                p.MainWindowTitle != "Microsoft PowerPoint" );

            return processes.Any();


            //// Check if Microsoft Word has any open documents
            //var processes = Process.GetProcessesByName("POWERPNT").Where(p => !string.IsNullOrEmpty(p.MainWindowTitle));

            //foreach (var process in processes)
            //{
            //    // Access the process main window title
            //    var mainWindowTitle = process.MainWindowTitle;

            //    // If a document is open, the main window title should not be empty
            //    if (!(mainWindowTitle == "PowerPoint" || mainWindowTitle == "Microsoft PowerPoint"))
            //    {
            //        return true;
            //    }
            //}

            //return false;
        }


        /// <summary>
        /// Gets the names of all open Presentations in Microsoft PowerPoint
        /// </summary>
        private static string[] GetPowerPointOpenPresentationNames()
        {
            // Retrieve the names of all open presentations in Microsoft PowerPoint
            var presentationNames = new ConcurrentBag<string>();

            var processes = Process.GetProcessesByName( "POWERPNT" )
                                   .Where( p => ! string.IsNullOrEmpty( p.MainWindowTitle ) );

            Parallel.ForEach(processes, process =>
                                        {
                                            // Access the process main window title and remove the " - Microsoft PowerPoint" or " - PowerPoint" suffix
                                            var mainWindowTitle = process.MainWindowTitle.Replace(" - Microsoft PowerPoint", "").Replace(" - PowerPoint", "");

                                            presentationNames.Add(mainWindowTitle);
                                        });

            return presentationNames.ToArray();
        }


        /// <summary>
        /// Checks if the Microsoft PowerPoint home screen is active
        /// </summary>
        private static bool IsHomeScreenActive()
        {
            var presentationNames = GetPowerPointOpenPresentationNames();

            if (presentationNames.Length <= 0) return false;

            var currentPresentationName = presentationNames[0];
            return !(currentPresentationName.EndsWith(" - Microsoft PowerPoint") || currentPresentationName.EndsWith(" - PowerPoint"));
        }





        /// <summary>
        /// Updates the presence
        /// </summary>
        private void UpdatePresence()
        {
            //Check if any presentation is open
            if (IsAnyOpenPresentation())
            {
                var presentationNames       = GetPowerPointOpenPresentationNames();
                var currentPresentationName = presentationNames[0];

                presence.UpdateDetails($"Editing presentation: {currentPresentationName}");
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