using System.Diagnostics;
using MBDRPC.Helpers;

namespace DRPCWorker
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                //if (_logger.IsEnabled(LogLevel.Information))
                //{
                //    _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
                //}


                CheckMicrosoftWord();
                CheckMicrosoftExcel();
                CheckMicrosoftPowerPoint();
                CheckMicrosoftAccess();
                CheckMicrosoftWhiteboard();
                CheckMicrosoftOneDrive();
                CheckMicrosoftPublisher();
                CheckMicrosoftOutlook();
                CheckMicrosoftPowerBI();


                await Task.Delay(1000, stoppingToken);
            }
        }




        private static  void CheckMicrosoftWord()
        {
            var appPath = "OfficeDRPC.exe";

            // Create a ProcessStartInfo object and specify the filename of the application to run
            var startInfo = new ProcessStartInfo
                            {
                                FileName               = appPath ,
                                UseShellExecute        = true ,
                                RedirectStandardOutput = true ,
                                RedirectStandardError  = true ,
                                CreateNoWindow         = true
                            };


            var isRunning = RunningAppChecker.IsAppRunning("winword");
            var isDRPCRunning = RunningAppChecker.IsAppRunning("OfficeDRPC");


            if (isRunning && !isDRPCRunning)
            {
                using (var process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.Start();
                }
            }
            else if (!isRunning && isDRPCRunning)
            {
                // Kill the process
                var process = Process.GetProcessesByName("OfficeDRPC");

                if (process.Length > 0 && !process[0].HasExited)
                {
                    process[0].Kill(); ;
                }
            }
        }
                 
        private static  void CheckMicrosoftExcel()
        {
            var appPath = "ExcelDRPC.exe";

            // Create a ProcessStartInfo object and specify the filename of the application to run
            var startInfo = new ProcessStartInfo
            {
                FileName = appPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                Verb = string.Empty
            };


            var isRunning = RunningAppChecker.IsAppRunning("excel");
            var isDRPCRunning = RunningAppChecker.IsAppRunning("ExcelDRPC");


            if (isRunning && !isDRPCRunning)
            {
                using (var process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.Start();
                }
            }
            else if (!isRunning && isDRPCRunning)
            {
                // Kill the process
                var process = Process.GetProcessesByName("ExcelDRPC");

                if (process.Length > 0 && !process[0].HasExited)
                {
                    process[0].Kill(); ;
                }
            }
        }
        
        public  static void CheckMicrosoftPowerPoint()
        {
            var appPath = "PowerPointDRPC.exe";

            // Create a ProcessStartInfo object and specify the filename of the application to run
            var startInfo = new ProcessStartInfo
            {
                FileName = appPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                Verb = string.Empty
            };


            var isRunning = RunningAppChecker.IsAppRunning("powerpnt");
            var isDRPCRunning = RunningAppChecker.IsAppRunningAndNotExited("PowerPointDRPC");


            if (isRunning && !isDRPCRunning)
            {
                using (var process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.Start();
                }
            }
            else if (!isRunning && isDRPCRunning)
            {
                // Kill the process
                var process = Process.GetProcessesByName("PowerPointDRPC");

                if (process.Length > 0 && !process[0].HasExited)
                {
                    process[0].Kill(); ;
                }
            }
        }
        
        public  static void CheckMicrosoftAccess()
        {
            var appPath = "AccessDRPC.exe";

            // Create a ProcessStartInfo object and specify the filename of the application to run
            var startInfo = new ProcessStartInfo
            {
                FileName = appPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                Verb = string.Empty
            };


            var isRunning = RunningAppChecker.IsAppRunning("msaccess");
            var isDRPCRunning = RunningAppChecker.IsAppRunningAndNotExited("AccessDRPC");


            if (isRunning && !isDRPCRunning)
            {
                using (var process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.Start();
                }
            }
            else if (!isRunning && isDRPCRunning)
            {
                // Kill the process
                var process = Process.GetProcessesByName("AccessDRPC");

                if (process.Length > 0 && !process[0].HasExited)
                {
                    process[0].Kill(); ;
                }
            }
        }
            
        public  static void CheckMicrosoftWhiteboard()
        {
            var appPath = "WhiteboardDRPC.exe";

            // Create a ProcessStartInfo object and specify the filename of the application to run
            var startInfo = new ProcessStartInfo
            {
                FileName = appPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                Verb = string.Empty
            };


            var isRunning = RunningAppChecker.IsAppRunning("MicrosoftWhiteboard");
            var isDRPCRunning = RunningAppChecker.IsAppRunningAndNotExited("WhiteboardDRPC");


            if (isRunning && !isDRPCRunning)
            {
                using (var process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.Start();
                }
            }
            else if (!isRunning && isDRPCRunning)
            {
                // Kill the process
                var process = Process.GetProcessesByName("WhiteboardDRPC");

                if (process.Length > 0 && !process[0].HasExited)
                {
                    process[0].Kill(); ;
                }
            }
        }
              
        public  static void CheckMicrosoftOneDrive()
        {
            var appPath = "OnedriveDRPC.exe";

            // Create a ProcessStartInfo object and specify the filename of the application to run
            var startInfo = new ProcessStartInfo
            {
                FileName = appPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                Verb = string.Empty
            };

            var isRunning = RunningAppChecker.IsAppRunning("onedrive");
            var isDRPCRunning = RunningAppChecker.IsAppRunningAndNotExited("OnedriveDRPC");


            if (isRunning && !isDRPCRunning)
            {
                using (var process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.Start();
                }
            }
            else if (!isRunning && isDRPCRunning)
            {
                // Kill the process
                var process = Process.GetProcessesByName("OnedriveDRPC");

                if (process.Length > 0 && !process[0].HasExited)
                {
                    process[0].Kill(); ;
                }
            }
        }
               
        public  static void CheckMicrosoftPublisher()
        {
            var appPath = "PublisherDRPC.exe";

            // Create a ProcessStartInfo object and specify the filename of the application to run
            var startInfo = new ProcessStartInfo
            {
                FileName = appPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                Verb = string.Empty
            };


            var isRunning = RunningAppChecker.IsAppRunning("mspub");
            var isDRPCRunning = RunningAppChecker.IsAppRunningAndNotExited("PublisherDRPC");


            if (isRunning && !isDRPCRunning)
            {
                using (var process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.Start();
                }
            }
            else if (!isRunning && isDRPCRunning)
            {
                // Kill the process
                var process = Process.GetProcessesByName("PublisherDRPC");

                if (process.Length > 0 && !process[0].HasExited)
                {
                    process[0].Kill(); ;
                }
            }
        }
              
        public  static void CheckMicrosoftOutlook()
        {
            const string appPath = "OutlookDRPC.exe";

            var isRunning = RunningAppChecker.IsAppRunning("olk");
            var isDRPCRunning = RunningAppChecker.IsAppRunningAndNotExited("OutlookDRPC");

            if (isRunning && !isDRPCRunning)
            {
                // Create a new process
                using (var process = new Process())
                {
                    // Set the process start info
                    process.StartInfo.FileName = appPath;

                    // Set options to hide the window
                    process.StartInfo.UseShellExecute = false;
                    process.StartInfo.CreateNoWindow = true;
                    process.StartInfo.RedirectStandardOutput = true;
                    process.StartInfo.RedirectStandardError = true;
                    process.StartInfo.Verb = string.Empty;

                    // Start the process
                    process.Start();
                }

            }
            else if (!isRunning && isDRPCRunning)
            {
                const string processName = "OutlookDRPC";
                var processes = Process.GetProcessesByName(processName);

                if (processes.Length > 0 && !processes[0].HasExited)
                {
                    processes[0].Kill();
                    processes[0].Close();
                    processes[0].Dispose();
                }
            }
        }
                 
        public  static void CheckMicrosoftPowerBI()
        {
            var appPath = "PowerBiDRPC.exe";

            // Create a ProcessStartInfo object and specify the filename of the application to run
            var startInfo = new ProcessStartInfo
            {
                FileName = appPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                Verb = string.Empty
            };


            var isRunning = RunningAppChecker.IsAppRunning("PBIDesktop");
            var isDRPCRunning = RunningAppChecker.IsAppRunningAndNotExited("PowerBiDRPC");


            if (isRunning && !isDRPCRunning)
            {
                using (var process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.Start();
                }
            }
            else if (!isRunning && isDRPCRunning)
            {
                // Kill the process
                var process = Process.GetProcessesByName("PowerBiDRPC");

                if (process.Length > 0 && !process[0].HasExited)
                {
                    process[0].Kill();
                }
            }
        }

    }

}
