namespace Aspose.Slides.Web.UI.Models.Interfaces
{
	public interface IMergerViewModel : IBaseViewModel
	{
		string StyleMasterUploadText { get; }

		IUploadFileModel StyleMasterUploadFile { get; }

		string ToFormat { get; }

		string[] Formats { get; }
	}
}
