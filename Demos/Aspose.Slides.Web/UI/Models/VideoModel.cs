using Aspose.Slides.Web.UI.Models.Interfaces;

namespace Aspose.Slides.Web.UI.Models
{
    class VideoModel : BaseViewModel, IVideoModel
    {
        public string Every { get; } = "Every slide";
        public string PageRange { get; } = "By slide range";
        public string RangeException { get; } = "Please, enter valid slide range in the field";
        public string[] VideoCodecs { get; } = { "MP4 (H264)", "MP4 (H265/HEVC)" };

    }
}