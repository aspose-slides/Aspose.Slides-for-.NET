namespace Aspose.Slides.Web.UI.Models.Interfaces
{
	public interface ISlideshowModel : IBaseViewModel
	{
		string FolderName { get; }

		string FileName { get; }

		string ProductName { get; }

		string FullscreenToggle { get; }
		string OverviewToggle { get; }
		string AutoplayTimerButton { get; }
		string OpenEditorButton { get; }
		string DownloadButton { get; }
	}
}
