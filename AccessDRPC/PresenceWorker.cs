using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

using AccessDRPC.Core;

using MBDRPC.Helpers;

namespace AccessDRPC
{
    public class PresenceWorker
    {
        private Presence presence                  = new Presence();
        private string   officeAppSubscriptionType = "Mirosoft Office";
        private bool     isFirstRun                = true;
        private DateTime startTime;
        private string   processName = "MSACCESS";

        public Timer Timer;



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
            Timer = new Timer(_ => CheckMicrosoftAccess(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
        }

        /// <summary>
        /// Stops the presence
        /// </summary>
        public void Stop()
        {
            presence.ShutDown();
            Timer.Dispose();
        }


        private void CheckMicrosoftAccess()
        {
            if (RunningAppChecker.IsMicrosoftAccessRunning())
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
            return Process
                  .GetProcessesByName("MSACCESS").Any(p => !string.IsNullOrEmpty(p.MainWindowTitle));
        }


        /// <summary>
        /// Gets the names of all open Databases/Windows in Microsoft Access
        /// </summary>
        private static string[] GetAccessOpenWindowNames()
        {
            // Retrieve the names of all open databases/windows in Microsoft Access
            var windowNames = Process.GetProcessesByName("MSACCESS")
                                     .Where(p => !string.IsNullOrEmpty(p.MainWindowTitle))
                                     .Select(process =>
                                              {
                                                  var mainWindowTitle = process.MainWindowTitle;

                                                  if (mainWindowTitle.Contains("-") && mainWindowTitle.Contains(":"))
                                                  {
                                                      var startIndex =
                                                          mainWindowTitle.IndexOf(" - ", StringComparison.Ordinal) + 2;
                                                      var endIndex =
                                                          mainWindowTitle.IndexOf(":", StringComparison.Ordinal);
                                                      var length = endIndex - startIndex - 1;

                                                      return mainWindowTitle.Substring(startIndex, length);
                                                  }

                                                  return mainWindowTitle;
                                              })
                                     .ToArray();

            return windowNames;
        }

        /// <summary>
        /// Checks if the Microsoft Access home screen window is active
        /// </summary>
        private static bool IsHomeScreenActive()
        {
            var handle = FindWindow(null, "Access");
            return handle != IntPtr.Zero;
        }


        /// <summary>
        /// Checks if the Microsoft Access home screen window is active
        /// </summary>
        private static bool IsHomeScreenActive(IReadOnlyList<string> openWindowNames)
        {
            if (openWindowNames.Count <= 0) return false;

            var windowName = openWindowNames[0];
            return windowName.Equals("Access");
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

            if ( windowTitle.Contains( "-" ) && windowTitle.Contains( ":" ) )
            {
                // Remove between 'Access - ' and ':' to the end from the window title
                var fileName = windowTitle.Substring( windowTitle.IndexOf( " - " , StringComparison.Ordinal ) + 3 ,
                                                      windowTitle.IndexOf( ":" ,   StringComparison.Ordinal ) -
                                                      windowTitle.IndexOf( " - " , StringComparison.Ordinal ) - 3 );

                return fileName;
            }

            return string.Empty;
        }



        /// <summary>
        /// Updates the presence
        /// </summary>
        private void UpdatePresence()
        {
            //Check if any database is open
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
                        presence.UpdateDetails($"Managing database: {activeWindowName}");
                    }
                }
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
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var office365Path = Path.Combine(appDataPath, "Microsoft", "Office");

            var programFilesPath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
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