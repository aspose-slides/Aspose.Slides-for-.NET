namespace Aspose.Slides.Web.UI.Models.Interfaces
{
	public interface IMessagesModel
	{
		string ServerErrorText { get; }
		string ProcessingTimeoutTitle { get; }
		string ProcessingTimeoutText { get; }
		string InvalidFileTitle { get; }
		string InvalidFileText { get; }
		string BadRequestTitle { get; }
		string BadRequestText { get; }
		string OtherErrorText { get; }
		string SuccessfullyUploaded { get; }
		string FileSelectMessage { get; }
		string ValidateEmailMessage { get; }
		string WrongRegExpMessage { get; }
		string NoSearchResultsMessage { get; }
		string UnlockInvalidPassword { get; }
	}
}
