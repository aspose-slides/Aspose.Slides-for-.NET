using Aspose.Slides.Web.UI.Models.Interfaces;

namespace Aspose.Slides.Web.UI.Models
{
    class SlideshowModel : BaseViewModel, ISlideshowModel
    {
        public string FolderName { get; set; }
        public string FileName { get; set; }
        public string ProductName { get; } = "Slides";
        public string FullscreenToggle { get; } = "Toggle full screen mode";
        public string OverviewToggle { get; } = "Toggle slides overview";
        public string AutoplayTimerButton { get; } = "Autoplay timer";
        public string OpenEditorButton { get; } = "Open in editor";
        public string DownloadButton { get; } = "Download";
    }
}