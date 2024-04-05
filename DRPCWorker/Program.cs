using System.Reflection;
using System.Runtime.InteropServices;
using DRPCWorker;
using Microsoft.Win32;

var builder = Host.CreateApplicationBuilder(args);

builder.Services.AddHostedService<Worker>();


var host = builder.Build();

host.Run();

// Register the app to be auto startup
using (var key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
{
    key?.SetValue("MBVRK.OfficeDRPCWorker", Assembly.GetExecutingAssembly().Location);
}
