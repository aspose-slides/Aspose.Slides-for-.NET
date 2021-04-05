namespace Aspose.Slides.Web.UI.Models.Interfaces
{
	public interface IEditorAppModel : IErrorReportModel, IMessagesModel, IPageViewModel
	{
		string FolderName { get; }
		string FileName { get; }
	}
}
