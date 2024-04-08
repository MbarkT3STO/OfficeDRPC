using System.Threading;
using MBDRPC.Helpers;
using WordWorker = OfficeDRPC.PresenceWorker;
using ExcelWorker = ExcelDRPC.PresenceWorker;
using PowerPointWorker = PowerPointDRPC.PresenceWorker;
using AccessWorker = AccessDRPC.PresenceWorker;
using PublisherWorker = PublisherDRPC.PresenceWorker;
using OutlookWorker = OutlookDRPC.PresenceWorker;
using WhiteboardWorker = WhiteboardDRPC.PresenceWorker;
using OneDriveWorker = OnedriveDRPC.PresenceWorker;
using PowerBiWorker = PowerBiDRPC.PresenceWorker;
using System;

namespace OfficeDRPCCommander
{
    public class Worker
    {
        public WordWorker       WordWorker       = new WordWorker();
        public ExcelWorker      ExcelWorker      = new ExcelWorker();
        public PowerPointWorker PowerPointWorker = new PowerPointWorker();
        public AccessWorker     AccessWorker     = new AccessWorker();
        public PublisherWorker  PublisherWorker  = new PublisherWorker();
        public OutlookWorker    OutlookWorker    = new OutlookWorker();
        public WhiteboardWorker WhiteboardWorker = new WhiteboardWorker();
        public OneDriveWorker   OneDriveWorker   = new OneDriveWorker();
        public PowerBiWorker    PowerBiWorker    = new PowerBiWorker();


        public Timer Timer;



        public void Start()
        {
            Timer = new Timer(_ => CheckDiscordAndStartWorker(), null, TimeSpan.Zero, TimeSpan.FromSeconds(30));

            //WordWorker.Start();
            //ExcelWorker.Start();
            //PowerPointWorker.Start();
            //AccessWorker.Start();
            //PublisherWorker.Start();
            //OutlookWorker.Start();
            //WhiteboardWorker.Start();
            //OneDriveWorker.Start();
            //PowerBiWorker.Start();
        }


        private void CheckDiscordAndStartWorker()
        {
            if (RunningAppChecker.IsDiscordRunning())
            {
                WordWorker.Start();
                ExcelWorker.Start();
                PowerPointWorker.Start();
                AccessWorker.Start();
                PublisherWorker.Start();
                OutlookWorker.Start();
                WhiteboardWorker.Start();
                OneDriveWorker.Start();
                PowerBiWorker.Start();
            }
            else
            {
                WordWorker.Timer.Dispose();
                ExcelWorker.Timer.Dispose();
                PowerPointWorker.Timer.Dispose();
                AccessWorker.Timer.Dispose();
                PublisherWorker.Timer.Dispose();
                OutlookWorker.Timer.Dispose();
                WhiteboardWorker.Timer.Dispose();
                OneDriveWorker.Timer.Dispose();
                PowerBiWorker.Timer.Dispose();
            }
        }
    }
}