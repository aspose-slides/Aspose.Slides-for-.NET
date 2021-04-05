namespace Aspose.Slides.Web.UI.Models.Interfaces
{
	public interface ISignatureModel : IBaseViewModel
	{
		string SelectFileText { get; }
		string SignatureDrawingError { get; }
		string SignatureTextError { get; }
		string SignatureTypeDrawing { get; }
		string SignatureTypeText { get; }
		string SignatureTypeImage { get; }
		string AddSignatureText { get; }
		string ToFormat { get; }
		string[] Formats { get; }
		IUploadFileModel UploadSignatureImage { get; }
	}
}
