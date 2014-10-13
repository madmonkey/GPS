namespace FileSystemWatch
{
	using System.IO;
	using System.Runtime.InteropServices;
	
	[GuidAttribute("B9E3551D-ACBD-431C-A60B-C637068066CF")]
	public enum ChangeType
	{
		Created = WatcherChangeTypes.Created,
		Deleted = WatcherChangeTypes.Deleted,
		Changed = WatcherChangeTypes.Changed,
		Renamed = WatcherChangeTypes.Renamed
	}
	[ClassInterface(ClassInterfaceType.None)]
	[GuidAttribute("6E0D7510-F2E9-4F94-969B-E867C1F0584D")]
	[ProgId("FileChangedEventArgs")]
	public class FileChangedEventArgs : IFileChangedEventArgs
	{
		public FileChangedEventArgs(FileSystemEventArgs args)
		{
			this.ChangeType = (ChangeType)args.ChangeType;
			this.FullPath = args.FullPath;
			this.Name = args.Name;

		}
		[DispId(0)]
		public ChangeType ChangeType { get; private set; }
		[DispId(1)]
		public string FullPath { get; private set; }
		[DispId(2)]
		public string Name { get; private set; }
	}
}
