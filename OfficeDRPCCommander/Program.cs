using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

using IWshRuntimeLibrary;

using Microsoft.Win32;
using static System.Net.Mime.MediaTypeNames;




namespace OfficeDRPCCommander
{
    internal class Program
    {

        [DllImport("kernel32.dll")]
        public static extern bool FreeConsole();


        static void Main(string[] args)
        {
            var worker = new Worker();

            try
            {
                // Hide the console window
                FreeConsole();

                var currentProcess = Process.GetCurrentProcess();
                currentProcess.PriorityClass = ProcessPriorityClass.RealTime;

                //// Register the app to be auto startup
                //string originalFilePath = Process.GetCurrentProcess().MainModule.FileName;
                //string shortcutPath = "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Startup";

                //// Create the shortcut
                // WshInterop.CreateShortcut(shortcutPath + "\\MBVRK.OfficeDRPC.lnk","OfficeDRPC",originalFilePath,"MBVRK.OfficeDRPC", "");


                // Path to the Startup folder
                string startupFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);

                // Path to your application executable
                string appExecutablePath = Process.GetCurrentProcess().MainModule.FileName;

                //// Create a WshShell object
                //var wshShell = new WshShell();

                //// Create a shortcut object
                //IWshShortcut shortcut = (IWshShortcut)wshShell.CreateShortcut(
                //    Path.Combine(startupFolderPath, "MBVRK.OfficeDRPC.lnk"));

                //// Set the target path of the shortcut
                //shortcut.TargetPath = appExecutablePath;

                //// Save the shortcut
                //shortcut.Save();


                //using (Microsoft.Win32.RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
                //{
                //    key.SetValue("MBVRK.OfficeDRPC", "\"" + appExecutablePath + "\"");
                //}



                // Remove the app from auto startup
                //using ( var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey( "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run" , true ) )
                //{
                //    key.DeleteValue( "MBVRK.OfficeDRPC" );
                //}


                worker.Start();

                currentProcess.PriorityClass = ProcessPriorityClass.BelowNormal;

                // Keep the app running
                Thread.Sleep(Timeout.Infinite);
            }
            catch (Exception ex)
            {
                // Create log file and write the exception message to it alongside with the app path
                var path = Path.Combine(Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName), "log.txt");

                // Open or create the log file and append text to it
                using (StreamWriter writer = new StreamWriter(path, true))
                {
                    writer.WriteLine(DateTime.Now.ToString("[yyyy-MM-dd HH:mm:ss] ") + ex.Message);
                }
            }
            finally
            {
                Environment.Exit(0);
            }
        }
    }
}
