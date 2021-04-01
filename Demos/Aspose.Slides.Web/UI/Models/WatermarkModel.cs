using Aspose.Slides.Web.UI.Models.Interfaces;

namespace Aspose.Slides.Web.UI.Models
{
    class WatermarkModel : BaseViewModel, IWatermarkModel
    {
        public IUploadFileModel ImageUploadFile { get; } = new UploadFileModel();
        public string RotateAngle { get; } = "Rotate Angle";
        public string Grayscaled { get; } = "Gray scaled";
        public string ZoomFactor { get; } = "Zoom factor";
        public string TextTitleSub { get; } = "Add text watermark in PowerPoint presentations files";
        public string ImageTitleSub { get; } = "Add image watermark in PowerPoint presentations files";
        public string AddedSuccessMessage { get; } = "Watermark added successfully";
    }
}