using System;
using System.IO;
using System.Threading;
using MBDRPC.Core;
using MBDRPC.Helpers;

namespace ExcelDRPC
{
	public class PresenceWorker
    {
        private string   officeAppSubscriptionType = "Mirosoft Office";
		private Presence presence                  = new Presence();
        private bool     isFirstRun                = true;
        private DateTime startTime;
        private string   processName;
        public  Timer    Timer;


        /// <summary>
        /// Starts the presence
        /// </summary>
        public void Start()
		{
            Timer = new Timer(_ => CheckMicrosoftExcel(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
        }


        private void CheckMicrosoftExcel()
        {
            var isMsExcelRunning = RunningAppChecker.IsAppRunning("excel");
            if (isMsExcelRunning)
            {
                if (isFirstRun)
                {
                    processName    = "EXCEL";
                    officeAppSubscriptionType = GetOfficeVersion();

                    presence.InitializePresence("1223982816459489350");

                    presence.UpdateLargeImage("exccellogo", "Microsoft Excel");
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