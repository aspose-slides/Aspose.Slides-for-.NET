namespace Aspose.Slides.Web.UI.Models.Interfaces
{
	public interface IBaseViewModel : IHeaderViewModel, IResultViewModel, IPageViewModel, IMessagesModel, IErrorReportModel
	{
		string App { get; }
		IUploadFileModel UploadFile { get; }
		string WorkButtonText { get; }
	}
}
