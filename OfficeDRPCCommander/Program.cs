using System;
using System.Diagnostics;
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

                var currentProcess = Process.GetCurrentProcess();
                currentProcess.PriorityClass = ProcessPriorityClass.BelowNormal;

                // Register the app to be auto startup
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
                {
                    key?.SetValue("MBVRK.OfficeDRPC", System.Reflection.Assembly.GetExecutingAssembly().Location);
                }

                // Remove the app from auto startup
                //using ( var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey( "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run" , true ) )
                //{
                //    key.DeleteValue( "MBVRK.OfficeDRPC" );
                //}


                worker.Start();

                // Keep the app running
                Thread.Sleep(Timeout.Infinite);
            }
            catch (Exception)
            {
                Environment.Exit(0);
            }

            Environment.Exit(0);
        }
    }
}
