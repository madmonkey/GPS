
namespace FileSystemWatch
{
	using System;
	using System.IO;
	using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Collections.Generic;

	[ClassInterface(ClassInterfaceType.None)]
	[GuidAttribute("D9CC7666-B1AC-4AA4-9905-C773A38236F7")]
	[ProgId("Watcher")]
	[ComSourceInterfaces(typeof(IWatcherEvents))]
	public class Watcher : IWatcher
    {
		[ComVisible(false)]
		public delegate void ChangeNotificationEventHandler(object sender, FileChangedEventArgs e);
		public event ChangeNotificationEventHandler OnChangeNotification;
        public event Action<string> LogData;

        private List<AliasData> aliasDataList;
        private FileSystemWatcher fsw;

        public Watcher()
        {
            aliasDataList = new List<AliasData>();

            UpdateAliasData();
        }
		
		public void BeginMonitoring(string path, string filter)
		{
			fsw = new FileSystemWatcher(path,filter)
			      	{
			      		EnableRaisingEvents = true,
						IncludeSubdirectories = true
			      	};
			fsw.Changed += TranslateEvent;
			fsw.Created += TranslateEvent;
			fsw.Deleted += TranslateEvent;
			fsw.Renamed += TranslateEvent;
		}

		public void StopMonitoring()
		{
			fsw.EndInit();
			fsw.Changed -= TranslateEvent;
			fsw.Created -= TranslateEvent;
			fsw.Deleted -= TranslateEvent;
			fsw.Renamed -= TranslateEvent;
		}

        public AliasData[] ObtainAliasDataCollection()
        {
            return aliasDataList.ToArray();
        }

        private void UpdateAliasData()
        {
            try
            {
                aliasDataList = AccessDatabaseLayer.BuildAlias("SELECT * FROM Alias");
                OnLogData("QVXR - UpdateAliasData: Count = " + aliasDataList.Count);
            }
            catch (Exception ex)
            {
                OnLogData("QVXR - Error Updating database: " + ex.Message);
            }
        }

        private void OnLogData(string data)
        {
            if (LogData != null)
                LogData(data);
        }

		private void TranslateEvent(object sender, FileSystemEventArgs e)
		{
            UpdateAliasData();
			var handler = OnChangeNotification;
			if (handler == null)
			{
				return;
			}

			foreach (ChangeNotificationEventHandler h in handler.GetInvocationList())
			{
				try
				{
					
					h(this, new FileChangedEventArgs(e));
				}
				catch (Exception)
				{
					Console.WriteLine("A listener threw an exception in its handler");
				}
			}
		}
	}
}
