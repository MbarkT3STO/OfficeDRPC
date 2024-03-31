using System;
using System.IO;
using System.Threading;
using MBDRPC.Core;
using MBDRPC.Helpers;

namespace OutlookDRPC
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
            Timer = new Timer(_ => CheckMicrosoftOutlook(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
        }


        private void CheckMicrosoftOutlook()
        {
            processName = "olk"; // Default value

            var isRunning = RunningAppChecker.IsOneAppRunningEndingWith( "Outlook" , "olk" );
            processName = RunningAppChecker.GetFirstProcessNameEndingWith( "Outlook" , "olk" );

            if (isRunning)
            {
                if (isFirstRun)
                {
                    officeAppSubscriptionType = GetOfficeVersion();

                    presence.InitializePresence("1224039960265359552");

                    presence.UpdateLargeImage("outlooklogo", "Microsoft Outlook");
                    presence.UpdateSmallImage("microsoft_365__2022_", GetOfficeVersion());

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
        /// Updates the presence
        /// </summary>
        private void UpdatePresence()
        {
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