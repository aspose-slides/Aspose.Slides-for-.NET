namespace Aspose.Slides.Web.UI.Models.Interfaces
{
	public interface IUploadFileModel
	{
		bool AcceptMultipleFiles { get; set; }
		string AcceptedExtensions { get; set; }
		string UploadId { get; }
		string Label { get; }
	}
}
