namespace FileSystemWatch
{
	using System.Runtime.InteropServices;

	[GuidAttribute("61A14EAA-4A22-4F86-9EA0-2F3490506071"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface IFileChangedEventArgs
	{
		[DispId(0)]
		ChangeType ChangeType { get; }

		[DispId(1)]
		string FullPath { get; }

		[DispId(2)]
		string Name { get; }
	}
}

