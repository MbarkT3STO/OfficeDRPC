using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using PowerBiDRPC.Core;

using MBDRPC.Helpers;

namespace PowerBiDRPC
{
    public class PresenceWorker
    {
        private          string   officeAppSubscriptionType = "Mirosoft Office";
        private readonly Presence presence                  = new Presence();
        private          bool     isFirstRun                = true;
        private          DateTime startTime;
        private const    string   processName = "PBIDesktop";

        public  Timer    Timer;



        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetWindowText(IntPtr hWnd, string lpString, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetClassName(IntPtr hWnd, [Out] StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        // Importing necessary function from kernel32.dll
        [DllImport("kernel32.dll")]
        private static extern IntPtr OpenProcess(uint dwDesiredAccess, bool bInheritHandle, uint dwProcessId);


        /// <summary>
        /// Gets the process ID associated with the window handle
        /// </summary>
        /// <param name="hWnd">Window handle</param>
        private static uint GetWindowProcessId(IntPtr hWnd)
        {
            GetWindowThreadProcessId(hWnd, out var processId);
            return processId;
        }

        /// <summary>
        /// Gets the name of the process associated with the process ID
        /// </summary>
        /// <param name="processId">Process ID</param>
        private static string GetProcessName(uint processId)
        {
            using (var process = Process.GetProcessById((int)processId))
            {
                return process.ProcessName;
            }
        }


        /// <summary>
        /// Starts the presence
        /// </summary>
        public void Start()
        {
            Timer = new Timer(_ => CheckMicrosoftPowerBI(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
        }


        /// <summary>
        /// Stops the presence
        /// </summary>
        public void Stop()
        {
            presence.ShutDown();
            Timer.Dispose();
        }


        private void CheckMicrosoftPowerBI()
        {
            if (RunningAppChecker.IsMicrosoftPowerBIRunning())
            {
                if (isFirstRun)
                {
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
            return Process
                  .GetProcessesByName("PBIDesktop").Any(p => !string.IsNullOrEmpty(p.MainWindowTitle));
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
            var handle = FindWindow(null, "Untitled - Power BI Desktop");
            return handle != IntPtr.Zero;
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
        /// Gets the name of the active window/file
        /// </summary>
        private static string GetActiveWindowName()
        {
            const string targetProcessName = "PBIDesktop";
            var          hWnd              = GetForegroundWindow();
            var          processId         = GetWindowProcessId(hWnd);
            var          processName       = GetProcessName(processId);

            if (!processName.Equals(targetProcessName, StringComparison.OrdinalIgnoreCase)) return string.Empty;

            const int nChars      = 256;
            var       windowTitle = new string(' ', nChars);

            return GetWindowText(hWnd, windowTitle, nChars) > 0 ? windowTitle : string.Empty;
        }


        /// <summary>
        /// Updates the presence
        /// </summary>
        private void UpdatePresence()
        {
            //Check if the home screen is active
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
            var office365Path = Path.Combine( appDataPath , "Microsoft" , "Office" );

            var programFilesPath    = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var perpetualOfficePath = Path.Combine( programFilesPath , "Microsoft Office" , "root" , "Office16" );

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