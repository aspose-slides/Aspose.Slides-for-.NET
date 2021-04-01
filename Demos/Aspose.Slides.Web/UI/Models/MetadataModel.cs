using Aspose.Slides.Web.UI.Models.Interfaces;

namespace Aspose.Slides.Web.UI.Models
{
    class MetadataModel : BaseViewModel, IMetadataModel
    {
        public string SaveButtonText { get; } = "Save";
        public string ClearButtonText { get; } = "Clear";
        public string CancelButtonText { get; } = "Cancel";
    }
}