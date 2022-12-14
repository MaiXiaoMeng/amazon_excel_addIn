using ExcelDna.Integration;
using ExcelDna.Integration.Extensibility;
using ExcelDna.IntelliSense;
using System;
using System.Runtime.InteropServices;

namespace AmazonExcelUDF
{
    [ComVisible(true)]
    [Guid("185734CB-10D4-4319-9983-AD5FD07441D5")]
    [ProgId("AmazonExcelUDF.Connection")]
    class ComAddInConnection : ExcelComAddIn
    {
        #region IDTExensibility2
        /// <summary>
        /// Receives notification that the Add-in is being loaded.
        /// </summary>
        public override void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            IntelliSenseServer.Install();
            base.OnConnection(Application, ConnectMode, AddInInst, ref custom);
        }
        /// <summary>
        /// Receives notification that the Add-in is being unloaded.
        /// </summary>
        public override void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            IntelliSenseServer.Uninstall();
            base.OnDisconnection(RemoveMode, ref custom);
        }
        /// <summary>
        /// Receives notification when the collection of Add-ins has changed.
        /// </summary>
        public override void OnAddInsUpdate(ref Array custom)
        {
            base.OnAddInsUpdate(ref custom);
        }
        /// <summary>
        /// Receives notification that the host application has completed loading.
        /// </summary>
        public override void OnStartupComplete(ref Array custom)
        {
            base.OnStartupComplete(ref custom);
        }
        /// <summary>
        /// Receives notification that the host application is being unloaded.
        /// </summary>
        public override void OnBeginShutdown(ref Array custom)
        {
            base.OnBeginShutdown(ref custom);
        }
        #endregion IDTExensibility2
    }
}
