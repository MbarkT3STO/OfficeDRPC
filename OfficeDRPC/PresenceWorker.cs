using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Threading;
using MBDRPC.Core;
using MBDRPC.Helpers;

using Microsoft.Win32;

namespace OfficeDRPC
{
	public class PresenceWorker
	{
		private Presence wordPresence   = new Presence();
        private bool     isWordFirstRun = true;
		private DateTime wordStartTime;
        private string   currentWordProcessName;
		private string   wordAppVersion;

        private string officeAppSubscriptionType = "Mirosoft Office";
        
		private Presence excelPresence  = new Presence();
        private bool     isExcelFirstRun = true;
        private DateTime excelStartTime;
        private string currentexcelProcessName;

        public Timer Timer;
        public Timer ExcelTimer;
        public Timer WorkerTimer;


        /// <summary>
        /// Starts the presence
        /// </summary>
        public void Start()
		{
            Timer = new Timer(_ => CheckMicrosoftWord(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
        }


        private void CheckMicrosoftWord()
        {
            var isMsWordRunning = RunningAppChecker.IsAppRunning("winword");
            if (isMsWordRunning)
            {
                if (isWordFirstRun)
                {
                    currentWordProcessName    = "WINWORD";
                    wordAppVersion            = "1.0.0.0";
                    officeAppSubscriptionType = GetOfficeVersion();

                    wordPresence.InitializePresence("1223964264449183765");

                    wordPresence.UpdateLargeImage("wordogo", "Microsoft Word");
                    wordPresence.UpdateSmallImage("microsoft_365__2022_", GetOfficeVersion());

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
        /// Updates the presence
        /// </summary>
        private void UpdateWordPresence()
		{
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