using WordWorker = OfficeDRPC.PresenceWorker;
using ExcelWorker = ExcelDRPC.PresenceWorker;
using PowerPointWorker = PowerPointDRPC.PresenceWorker;
using AccessWorker = AccessDRPC.PresenceWorker;
using PublisherWorker = PublisherDRPC.PresenceWorker;
using OutlookWorker = OutlookDRPC.PresenceWorker;
using WhiteboardWorker = WhiteboardDRPC.PresenceWorker;
using OneDriveWorker = OnedriveDRPC.PresenceWorker;
using PowerBiWorker = PowerBiDRPC.PresenceWorker;

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



        public void Start()
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
    }
}