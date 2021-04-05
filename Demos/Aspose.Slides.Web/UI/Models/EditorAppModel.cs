using Aspose.Slides.Web.UI.Models.Interfaces;

namespace Aspose.Slides.Web.UI.Models
{
    class EditorAppModel : BaseViewModel, IEditorAppModel
    {
        public string FolderName { get; set; }
        public string FileName { get; set; }
    }
}