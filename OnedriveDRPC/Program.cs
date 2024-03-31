using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OnedriveDRPC;

namespace OnedriveDRPC
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Hide the console window
            //var handle = GetConsoleWindow();
            //// To hide:
            //ShowWindow(handle, SW_HIDE);

            // Register the app to be auto startup
            //using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
            //{
            //    key?.SetValue("MBVRK.OfficeDRPC", System.Reflection.Assembly.GetExecutingAssembly().Location);
            //}

            var worker = new PresenceWorker();

            worker.Start();


            Console.ReadKey();

            // Dispose the timers
            //worker.WordTimer.Dispose();
            worker.Timer.Dispose();

            // Kill the application
            Environment.Exit(0);

            //    try
            //    {
            //        worker.Start();

            //        // sleep thread forever
            //        Thread.Sleep(Timeout.Infinite);
            //    }
            //    catch (Exception)
            //    {
            //        // Dispose the timer
            //        worker.Timer.Dispose();

            //        // Kill the application
            //        Environment.Exit(0);
            //    }
        }
    }
}
