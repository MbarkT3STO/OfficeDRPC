using System;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using MBDRPC.Core;
using MBDRPC.Helpers;

namespace OfficeDRPC
{
    public class PresenceWorker
    {
        private Presence wordPresence   = new Presence();
        private bool     isWordFirstRun = true;
        private DateTime wordStartTime;
        private string   currentWordProcessName;
        private string   officeAppSubscriptionType = "Mirosoft Office";
        public  Timer    Timer;
  

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
                    currentWordProcessName = "WINWORD";
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
        /// Checks if any Microsoft Word documents are open
        /// </summary>
        private static bool IsAnyWordDocumentOpen()
        {
            // Check if Microsoft Word has any open documents
            var processes = Process.GetProcessesByName( "WINWORD" )
                                   .Where( p => ! string.IsNullOrEmpty( p.MainWindowTitle ) &&
                                                p.MainWindowTitle != "Word"                 &&
                                                p.MainWindowTitle != "Microsoft Word" );

            return processes.Any();
        }


        /// <summary>
        /// Gets the names of all open documents in Microsoft Word
        /// </summary>
        private static string[] GetWordOpenDocumentsNames()
        {
            // Retrieve the names of all open documents in Microsoft Word
            var documentNames = new ConcurrentBag<string>();

            var processes = Process.GetProcessesByName( "WINWORD" )
                                   .Where( p => ! string.IsNullOrEmpty( p.MainWindowTitle ) );

            Parallel.ForEach(processes, process =>
                                        {
                                            // Access the process main window title and remove the " - Microsoft Word" or " - Word" suffix
                                            var mainWindowTitle = process.MainWindowTitle.Replace(" - Microsoft Word", "").Replace(" - Word", "");

                                            documentNames.Add( mainWindowTitle );
                                        });

            return documentNames.ToArray();
        }


        /// <summary>
        /// Checks if the Microsoft Word home screen is active
        /// </summary>
        private static bool IsHomeScreenActive()
        {
            var documentNames = GetWordOpenDocumentsNames();

            if ( documentNames.Length <= 0 ) return false;

            var documentName = documentNames[0];
            return !(documentName.EndsWith(" - Microsoft Word") || documentName.EndsWith(" - Word"));
        }






        /// <summary>
        /// Updates the presence
        /// </summary>
        private void UpdateWordPresence()
        {
            //Check if any documents are open
            if (IsAnyWordDocumentOpen())
            {
                var documentNames = GetWordOpenDocumentsNames();
                var currentDocumentName = documentNames[0].Replace( " - Microsoft Word" , "" )
                                                          .Replace( " - Word" , "" );

                wordPresence.UpdateDetails($"Editing document: {currentDocumentName}");
            }
            else if (IsHomeScreenActive())
            {
                wordPresence.UpdateDetails("Home screen");
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