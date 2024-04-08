using System;
using System.IO;
using System.Threading;

using OnedriveDRPC.Core;

using MBDRPC.Helpers;

namespace OnedriveDRPC
{
	public class PresenceWorker
    {
        private readonly Presence presence                  = new Presence();
        private          string   officeAppSubscriptionType = "Mirosoft Office";
        private          bool     isFirstRun                = true;
        private          DateTime startTime;
        private const    string   processName = "OneDrive";

        public Timer Timer;


        /// <summary>
        /// Starts the presence
        /// </summary>
        public void Start()
		{
            Timer = new Timer(_ => CheckMicrosoftOnedrive(), null, TimeSpan.Zero, TimeSpan.FromSeconds(20));
        }


        private void CheckMicrosoftOnedrive()
        {
            if (RunningAppChecker.IsMicrosoftOneDriveRunning())
            {
                if (isFirstRun)
                {
                    officeAppSubscriptionType = GetOfficeVersion();

                    presence.InitializePresence("1224035031597453312");

                    presence.UpdateLargeImage("onedrivelogo", "Microsoft OneDrive");
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
            //UpdatePresenceTime();
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