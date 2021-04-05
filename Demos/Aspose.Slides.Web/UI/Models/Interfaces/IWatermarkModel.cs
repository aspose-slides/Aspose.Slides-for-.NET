namespace Aspose.Slides.Web.UI.Models.Interfaces
{
	public interface IWatermarkModel : IBaseViewModel
	{
		IUploadFileModel ImageUploadFile { get; }
		string RotateAngle { get; }
		string Grayscaled { get; }
		string ZoomFactor { get; }
		string TextTitleSub { get; }
		string ImageTitleSub { get; }
		string AddedSuccessMessage { get; }
	}
}
