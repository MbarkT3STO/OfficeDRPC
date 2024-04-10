using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

using MBDRPC.Helpers;

using OfficeDRPC.Core;

namespace OfficeDRPC
{
    public class PresenceWorker
    {
        private Presence wordPresence   = new Presence();
        private bool     isWordFirstRun = true;
        private DateTime wordStartTime;
        private string   currentWordProcessName    = "WINWORD";
        private string   officeAppSubscriptionType = "Mirosoft Office";

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



        /// <summary>
        /// Starts the presence
        /// </summary>
        public void Start()
        {
            Timer = new Timer(_ => CheckMicrosoftWord(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
        }


        /// <summary>
        /// Stops the presence
        /// </summary>
        public void Stop()
        {
            wordPresence.ShutDown();
            Timer.Dispose();
        }


        private void CheckMicrosoftWord()
        {
            if (RunningAppChecker.IsMicrosoftWordRunning())
            {
                if (isWordFirstRun)
                {
                    officeAppSubscriptionType = GetOfficeVersion();

                    wordPresence.InitializePresence("1223964264449183765");

                    wordPresence.UpdateLargeImage("wordogo", "Microsoft Word");
                    wordPresence.UpdateSmallImage("microsoft_365__2022_", officeAppSubscriptionType);

                    wordPresence.UpdateDetails(officeAppSubscriptionType);


                    wordStartTime = RunningAppChecker.GetProcessStartTime(currentWordProcessName);
                    isWordFirstRun = false;
                }

                UpdateWordPresence();
            }
            else
            {
                wordPresence.ShutDown();
                isWordFirstRun = true;
            }
        }



        /// <summary>
        /// Checks if any Microsoft Word documents/Windows are open
        /// </summary>
        private static bool IsAnyWordWindowOpen()
        {
            // Check if Microsoft Word has any open documents
            return Process
                  .GetProcessesByName("WINWORD").Any(p => !string.IsNullOrEmpty(p.MainWindowTitle));
        }


        /// <summary>
        /// Gets the names of all open documents/windows in Microsoft Word
        /// </summary>
        private static string[] GetWordOpenWindowNames()
        {
            // Retrieve the names of all open documents/windows in Microsoft Word
            var windowNames = Process.GetProcessesByName("WINWORD")
                                     .Where(p => !string.IsNullOrEmpty(p.MainWindowTitle))
                                     .Select(p => p.MainWindowTitle.Replace(" - Word", ""))
                                     .ToArray();

            return windowNames;
        }


        /// <summary>
        /// Checks if the Microsoft Word home screen window is active
        /// </summary>
        private static bool IsHomeScreenActive()
        {
            var handle = FindWindow(null, "Word");
            return handle != IntPtr.Zero;
        }


        /// <summary>
        /// Checks if the Microsoft Word home screen window is active
        /// </summary>
        private static bool IsHomeScreenActive(IReadOnlyList<string> openWindowNames)
        {
            if (openWindowNames.Count <= 0) return false;

            var windowName = openWindowNames[0];
            return windowName.Equals("Word");
        }


        /// <summary>
        /// Gets the name of the active window/file
        /// </summary>
        private static string GetActiveWindowName()
        {
            // Microsoft  Word is running, check for the active window
            var foregroundWindow = GetForegroundWindow();

            if (foregroundWindow == IntPtr.Zero) return string.Empty;

            // Get the window title
            const int nChars = 256;
            var windowTitle = new string(' ', nChars);
            GetWindowText(foregroundWindow, windowTitle, nChars);

            var trimmedWindowTitle = windowTitle.Trim().Trim("\0".ToCharArray());

            if (!trimmedWindowTitle.EndsWith("- Word", StringComparison.Ordinal))
                return string.Empty;

            // Remove from ' - Word' to the end from the window title
            var fileName = windowTitle.Substring(0, windowTitle.IndexOf(" - Word", StringComparison.Ordinal));

            return fileName;

        }



        /// <summary>
        /// Updates the presence
        /// </summary>
        private void UpdateWordPresence()
        {
            //Check if any documents are open
            if (IsAnyWordWindowOpen())
            {
                if (IsHomeScreenActive())
                {
                    wordPresence.UpdateDetails("Home screen");
                }
                else
                {
                    var activeWindowName = GetActiveWindowName();

                    if (activeWindowName != string.Empty)
                    {
                        wordPresence.UpdateDetails($"Editing: {activeWindowName}");
                    }
                }
            }
            else
            {
                wordPresence.UpdateDetails(officeAppSubscriptionType);
            }

            UpdateWordPresenceTime();
            wordPresence.UpdatePresence();
        }


        /// <summary>
        /// Updates the presence time
        /// </summary>
        private void UpdateWordPresenceTime()
        {
            var elapsedTime = (DateTime.Now - wordStartTime).ToString(@"hh\:mm\:ss");
            wordPresence.UpdateState(elapsedTime);
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