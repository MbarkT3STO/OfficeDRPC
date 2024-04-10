using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
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
        private string   processName = "MSPUB";

        public  Timer    Timer;




        [DllImport( "user32.dll", SetLastError = true)]
        static extern IntPtr GetForegroundWindow();

        [DllImport( "user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport( "user32.dll", SetLastError = true)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport( "user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetWindowText(IntPtr hWnd, string lpString, int nMaxCount);

        [DllImport( "user32.dll", SetLastError = true)]
        static extern int GetClassName(IntPtr hWnd, [Out] StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);




        /// <summary>
        /// Starts the presence
        /// </summary>
        public void Start()
		{
            Timer = new Timer(_ => CheckMicrosoftPublisher(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
        }

        /// <summary>
        /// Stops the presence
        /// </summary>
        public void Stop()
        {
            presence.ShutDown();
            Timer.Dispose();
        }


        private void CheckMicrosoftPublisher()
        {
            if (RunningAppChecker.IsMicrosoftPublisherRunning())
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
            return Process
                  .GetProcessesByName("MSPUB").Any(p => !string.IsNullOrEmpty(p.MainWindowTitle));
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
            var handle = FindWindow(null, "Publisher");
            return handle != IntPtr.Zero;
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
        /// Gets the name of the active window/file
        /// </summary>
        private static string GetActiveWindowName()
        {
            // App is running, check for the active window
            var foregroundWindow = GetForegroundWindow();

            if (foregroundWindow == IntPtr.Zero) return string.Empty;

            // Get the window title
            const int nChars      = 256;
            var       windowTitle = new string(' ', nChars);
            GetWindowText(foregroundWindow, windowTitle, nChars);

            var trimmedWindowTitle = windowTitle.Trim().Trim("\0".ToCharArray());

            if (!trimmedWindowTitle.EndsWith("- Publisher", StringComparison.Ordinal))
                return string.Empty;

            // Remove from ' - Publisher' to the end from the window title
            var fileName = windowTitle.Substring(0, windowTitle.IndexOf(" - Publisher", StringComparison.Ordinal));

            return fileName;

        }





        /// <summary>
        /// Updates the presence
        /// </summary>
        private void UpdatePresence()
        {
            //Check if any composition is open
            if (IsAnyOpenWindow())
            {
                if (IsHomeScreenActive())
                {
                    presence.UpdateDetails("Home screen");
                }
                else
                {
                    var activeWindowName = GetActiveWindowName();

                    if (activeWindowName != string.Empty)
                    {
                        presence.UpdateDetails($"Editing: {activeWindowName}");
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
            var appDataPath   = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var office365Path = Path.Combine(appDataPath, "Microsoft", "Office");

            var programFilesPath    = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var perpetualOfficePath = Path.Combine(programFilesPath, "Microsoft Office", "root", "Office16");

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