
namespace FileSystemWatch
{
	using System.Runtime.InteropServices;

	[GuidAttribute("5487C12F-E9AD-41F7-90FF-1F6EA3BB93AD"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface IWatcher
	{
		[DispId(0)]
		void BeginMonitoring(string path, string filter);

		[DispId(1)]
		void StopMonitoring();

        [DispId(2)]
        AliasData[] ObtainAliasDataCollection();
		
	}
	[GuidAttribute("8B287E61-F94D-47E0-8C43-BC1DF711EA21"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
	public interface IWatcherEvents
	{
		[DispId(0)]
		void OnChangeNotification(object sender, FileChangedEventArgs fileChangedEventArgs );
        [DispId(1)]
        void LogData(string data);
	}
}

