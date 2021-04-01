using Aspose.Slides.Web.UI.Models.Interfaces;

namespace Aspose.Slides.Web.UI.Models
{
    class RedactionModel : BaseViewModel, IRedactionModel
    {
        public string SearchQuery { get; } = "Type text or regular expression";
        public string ReplaceText { get; } = "Type replace text";
    }
}