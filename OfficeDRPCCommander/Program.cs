using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace OfficeDRPCCommander
{
    internal class Program
    {
        [DllImport( "kernel32.dll")]
        public static extern bool FreeConsole();

        static void Main(string[] args)
        {
            var worker = new Worker();

            try
            {
                // Hide the console window
                FreeConsole();

                // Register the app to be auto startup
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
                {
                    key?.SetValue("MBVRK.OfficeDRPC", System.Reflection.Assembly.GetExecutingAssembly().Location);
                }


                worker.Start();

                // Keep the app running
                Thread.Sleep(Timeout.Infinite);
            }
            catch (Exception)
            {
                worker.WordTimer.Dispose();
                worker.ExcelTimer.Dispose();
                worker.PowerPointTimer.Dispose();
                worker.AccessTimer.Dispose();
                worker.WhiteboardTimer.Dispose();
                worker.OneDriveTimer.Dispose();
                worker.PublisherTimer.Dispose();
                worker.OutlookTimer.Dispose();

                Environment.Exit(0);
            }

            //Console.ReadKey();

            //worker.WordTimer.Dispose();
            //worker.ExcelTimer.Dispose();
            //worker.PowerPointTimer.Dispose();
            //worker.AccessTimer.Dispose();
            //worker.WhiteboardTimer.Dispose();
            //worker.OneDriveTimer.Dispose();
            //worker.PublisherTimer.Dispose();
            //worker.OutlookTimer.Dispose();

            //Environment.Exit(0);
        }
    }
}
