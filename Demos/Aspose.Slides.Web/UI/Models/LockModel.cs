using Aspose.Slides.Web.UI.Models.Interfaces;

namespace Aspose.Slides.Web.UI.Models
{
    class LockModel : BaseViewModel, ILockModel
    {
        public string EditPassword { get; } = "Password for edit protection";
        public string ViewPassword { get; } = "Password for view protection";
        public string MarkAsFinal { get; } = "Mark as final";
    }
}