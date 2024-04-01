using System;
using System.Diagnostics;
using System.IO;
using System.Threading;

using MBDRPC.Helpers;

namespace OfficeDRPCCommander
{
    public class Worker
    {
        public Timer WordTimer;
        public Timer ExcelTimer;
        public Timer PowerPointTimer;
        public Timer AccessTimer;
        public Timer WhiteboardTimer;
        public Timer OneDriveTimer;
        public Timer PublisherTimer;
        public Timer OutlookTimer;


        public void Start()
        {
            WordTimer = new Timer(_ => CheckMicrosoftWord(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
            ExcelTimer = new Timer(_ => CheckMicrosoftExcel(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
            PowerPointTimer = new Timer(_ => CheckMicrosoftPowerPoint(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
            AccessTimer = new Timer(_ => CheckMicrosoftAccess(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
            WhiteboardTimer = new Timer(_ => CheckMicrosoftWhiteboard(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
            OneDriveTimer = new Timer(_ => CheckMicrosoftOneDrive(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
            PublisherTimer = new Timer(_ => CheckMicrosoftPublisher(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
            OutlookTimer = new Timer(_ => CheckMicrosoftOutlook(), null, TimeSpan.Zero, TimeSpan.FromSeconds(1));
        }


        private void CheckMicrosoftWord()
        {
            //var appPath = "C:\\program files\\MBVRK\\OfficeDRPC\\WordDRPC\\OfficeDRPC.exe";
            //const string appPath = @"C:\Users\MBARK\source\repos\MbarkT3STO\OfficeDRPC\OfficeDRPC\bin\Debug\OfficeDRPC.exe";
            const string appPath = @"C:\Users\MBARK.AzureAD\source\repos\MbarkT3STO\OfficeDRPC\OfficeDRPC\bin\Debug\OfficeDRPC.exe";

            // Create a ProcessStartInfo object and specify the filename of the application to run
            var startInfo = new ProcessStartInfo
            {
                FileName = appPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };


            var isRunning = RunningAppChecker.IsAppRunning("winword");
            var isDRPCRunning = RunningAppChecker.IsAppRunning("OfficeDRPC");


            if (isRunning && !isDRPCRunning)
            {
                using (var process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.Start();
                    process.WaitForExit();
                }
            }
            else if (!isRunning && isDRPCRunning)
            {
                // Kill the process
                var process = Process.GetProcessesByName("OfficeDRPC");

                if (process.Length > 0)
                {
                    process[0].Kill();
                }
            }
        }


        private void CheckMicrosoftExcel()
        {
            //var appPath = "C:\\program files\\MBVRK\\OfficeDRPC\\WordDRPC\\ExcelDRPC.exe";
            //const string appPath = @"C:\Users\MBARK\source\repos\MbarkT3STO\OfficeDRPC\ExcelDRPC\bin\Debug\ExcelDRPC.exe";
            const string appPath = @"C:\Users\MBARK.AzureAD\source\repos\MbarkT3STO\OfficeDRPC\ExcelDRPC\bin\Debug\ExcelDRPC.exe";

            // Create a ProcessStartInfo object and specify the filename of the application to run
            var startInfo = new ProcessStartInfo
            {
                FileName = appPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };


            var isRunning = RunningAppChecker.IsAppRunning("excel");
            var isDRPCRunning = RunningAppChecker.IsAppRunning("ExcelDRPC");


            if (isRunning && !isDRPCRunning)
            {
                using (var process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.Start();
                    process.WaitForExit();
                }
            }
            else if (!isRunning && isDRPCRunning)
            {
                // Kill the process
                var process = Process.GetProcessesByName("ExcelDRPC");

                if (process.Length > 0)
                {
                    process[0].Kill();
                }
            }
        }


        public void CheckMicrosoftPowerPoint()
        {
            //var appPath = "C:\\program files\\MBVRK\\OfficeDRPC\\PowerPointDRPC\\PowerPointDRPC.exe";
            //const string appPath = @"C:\Users\MBARK\source\repos\MbarkT3STO\OfficeDRPC\PowerPointDRPC\bin\Debug\PowerPointDRPC.exe";
            const string appPath = @"C:\Users\MBARK.AzureAD\source\repos\MbarkT3STO\OfficeDRPC\PowerPointDRPC\bin\Debug\PowerPointDRPC.exe";

            // Create a ProcessStartInfo object and specify the filename of the application to run
            var startInfo = new ProcessStartInfo
            {
                FileName = appPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };


            var isRunning = RunningAppChecker.IsAppRunning("powerpnt");
            var isDRPCRunning = RunningAppChecker.IsAppRunning("PowerPointDRPC");


            if (isRunning && !isDRPCRunning)
            {
                using (var process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.Start();
                    process.WaitForExit();
                }
            }
            else if (!isRunning && isDRPCRunning)
            {
                // Kill the process
                var process = Process.GetProcessesByName("PowerPointDRPC");

                if (process.Length > 0)
                {
                    process[0].Kill();
                }
            }
        }


        public void CheckMicrosoftAccess()
        {
            //var appPath = "C:\\program files\\MBVRK\\OfficeDRPC\\AccessDRPC\\AccessDRPC.exe";
            //const string appPath = @"C:\Users\MBARK\source\repos\MbarkT3STO\OfficeDRPC\AccessDRPC\bin\Debug\AccessDRPC.exe";
            const string appPath = @"C:\Users\MBARK.AzureAD\source\repos\MbarkT3STO\OfficeDRPC\AccessDRPC\bin\Debug\AccessDRPC.exe";

            // Create a ProcessStartInfo object and specify the filename of the application to run
            var startInfo = new ProcessStartInfo
            {
                FileName = appPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };


            var isRunning = RunningAppChecker.IsAppRunning("msaccess");
            var isDRPCRunning = RunningAppChecker.IsAppRunning("AccessDRPC");


            if (isRunning && !isDRPCRunning)
            {
                using (var process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.Start();
                    process.WaitForExit();
                }
            }
            else if (!isRunning && isDRPCRunning)
            {
                // Kill the process
                var process = Process.GetProcessesByName("AccessDRPC");

                if (process.Length > 0)
                {
                    process[0].Kill();
                }
            }
        }


        public void CheckMicrosoftWhiteboard()
        {
            //var appPath = "C:\\program files\\MBVRK\\OfficeDRPC\\WhiteboardDRPC\\WhiteboardDRPC.exe";
            //const string appPath = @"C:\Users\MBARK\source\repos\MbarkT3STO\OfficeDRPC\WhiteboardDRPC\bin\Debug\WhiteboardDRPC.exe";
            const string appPath = @"C:\Users\MBARK.AzureAD\source\repos\MbarkT3STO\OfficeDRPC\WhiteboardDRPC\bin\Debug\WhiteboardDRPC.exe";

            // Create a ProcessStartInfo object and specify the filename of the application to run
            var startInfo = new ProcessStartInfo
            {
                FileName = appPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };


            var isRunning = RunningAppChecker.IsAppRunning("MicrosoftWhiteboard");
            var isDRPCRunning = RunningAppChecker.IsAppRunning("WhiteboardDRPC");


            if (isRunning && !isDRPCRunning)
            {
                using (var process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.Start();
                    process.WaitForExit();
                }
            }
            else if (!isRunning && isDRPCRunning)
            {
                // Kill the process
                var process = Process.GetProcessesByName("WhiteboardDRPC");

                if (process.Length > 0)
                {
                    process[0].Kill();
                }
            }
        }


        public void CheckMicrosoftOneDrive()
        {
            //var appPath = "C:\\program files\\MBVRK\\OfficeDRPC\\OnedriveDRPC\\OnedriveDRPC.exe";
            //const string appPath = @"C:\Users\MBARK\source\repos\MbarkT3STO\OfficeDRPC\OnedriveDRPC\bin\Debug\OnedriveDRPC.exe";
            const string appPath = @"C:\Users\MBARK.AzureAD\source\repos\MbarkT3STO\OfficeDRPC\OnedriveDRPC\bin\Debug\OnedriveDRPC.exe";

            // Create a ProcessStartInfo object and specify the filename of the application to run
            var startInfo = new ProcessStartInfo
            {
                FileName = appPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };

            var isRunning = RunningAppChecker.IsAppRunning("onedrive");
            var isDRPCRunning = RunningAppChecker.IsAppRunning("OnedriveDRPC");


            if (isRunning && !isDRPCRunning)
            {
                using (var process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.Start();
                    process.WaitForExit();
                }
            }
            else if (!isRunning && isDRPCRunning)
            {
                // Kill the process
                var process = Process.GetProcessesByName("OnedriveDRPC");

                if (process.Length > 0)
                {
                    process[0].Kill();
                }
            }
        }


        public void CheckMicrosoftPublisher()
        {
            //var appPath = "C:\\program files\\MBVRK\\OfficeDRPC\\PublisherDRPC\\PublisherDRPC.exe";
            //const string appPath = @"C:\Users\MBARK\source\repos\MbarkT3STO\OfficeDRPC\PublisherDRPC\bin\Debug\PublisherDRPC.exe";
            const string appPath = @"C:\Users\MBARK.AzureAD\source\repos\MbarkT3STO\OfficeDRPC\PublisherDRPC\bin\Debug\PublisherDRPC.exe";

            // Create a ProcessStartInfo object and specify the filename of the application to run
            var startInfo = new ProcessStartInfo
            {
                FileName = appPath,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };


            var isRunning = RunningAppChecker.IsAppRunning("mspub");
            var isDRPCRunning = RunningAppChecker.IsAppRunning("PublisherDRPC");


            if (isRunning && !isDRPCRunning)
            {
                using (var process = new Process())
                {
                    process.StartInfo = startInfo;
                    process.Start();
                    process.WaitForExit();
                }
            }
            else if (!isRunning && isDRPCRunning)
            {
                // Kill the process
                var process = Process.GetProcessesByName("PublisherDRPC");

                if (process.Length > 0)
                {
                    process[0].Kill();
                }
            }
        }


        public void CheckMicrosoftOutlook()
        {
            ////var appPath = "C:\\program files\\MBVRK\\OfficeDRPC\\OutlookDRPC\\OutlookDRPC.exe";
            //const string appPath = @"C:\Users\MBARK\source\repos\OfficeDRPC\OutlookDRPC\bin\Debug\OutlookDRPC.exe";

            //// Create a ProcessStartInfo object and specify the filename of the application to run
            //var startInfo = new ProcessStartInfo
            //                {
            //                    FileName               = appPath,
            //                    UseShellExecute        = false,
            //                    RedirectStandardOutput = true,
            //                    CreateNoWindow         = true
            //                };

            //var isRunning     = RunningAppChecker.IsAppRunning("olk");
            //var isDRPCRunning = RunningAppChecker.IsAppRunning("OutlookDRPC");


            //if (isRunning && !isDRPCRunning)
            //{
            //    using (var process = new Process())
            //    {
            //        process.StartInfo = startInfo;
            //        process.Start();
            //        process.WaitForExit();
            //    }
            //}
            //else if (!isRunning && isDRPCRunning)
            //{
            //    // Kill the process
            //    var process = Process.GetProcessesByName("OutlookDRPC");

            //    foreach ( var process1 in process )
            //    {
            //        process1.Kill();
            //    }
            //}


            //const string appPath = @"C:\Users\MBARK\source\repos\MbarkT3STO\OfficeDRPC\OutlookDRPC\bin\Debug\OutlookDRPC.exe";
            const string appPath = @"C:\Users\MBARK.AzureAD\source\repos\MbarkT3STO\OfficeDRPC\OutlookDRPC\bin\Debug\OutlookDRPC.exe";
            var isRunning = RunningAppChecker.IsAppRunning("olk");
            var isDRPCRunning = RunningAppChecker.IsAppRunning( "OutlookDRPC" );

            if (isRunning && !isDRPCRunning)
            {
                // Start the process with no window
                Process.Start(appPath);
            }
            else if (!isRunning && isDRPCRunning)
            {
                var processName = "OutlookDRPC";
                var processes = Process.GetProcessesByName(processName);
                processes[0].Kill();
            }
        }
    }
}