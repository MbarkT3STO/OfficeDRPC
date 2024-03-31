using System;

namespace OfficeDRPCCommander
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var worker = new Worker();

            worker.Start();

            Console.ReadKey();

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
    }
}
