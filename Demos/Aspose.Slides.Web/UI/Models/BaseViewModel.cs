using Aspose.Slides.Web.UI.Models.Interfaces;

namespace Aspose.Slides.Web.UI.Models
{
    class BaseViewModel : IBaseViewModel
    {
        public string ProductTitle { get; set; }
        public string ProductTitleSub { get; set; }
        public string SuccessMessage { get; }
        public string DownloadButtonText { get; } = "Download";
        public string ViewButtonText { get; } = "View";
        public string EditButtonText { get; } = "Edit";
        public string BookmarkText { get; }
        public string ProcessAnotherText { get; } = "Process another presentation";
        public string CloudApiLink { get; } = "https://products.aspose.cloud/slides/family";
        public string OnPremiseApiLink { get; } = "https://products.aspose.com/slides/family";
        public string SendToText { get; } = "Send";
        public string EmailTo { get; }
        public string PageTitle { get; } = "Demo Slides App";
        public string MetaDescription { get; } = "Demo Slides App";
        public string APIBasePath { get; set; }
        public string App { get; } = "Demo app";
        public IUploadFileModel UploadFile { get; } = new UploadFileModel();
        public string WorkButtonText { get; }

        public string ServerErrorText { get; } = "Server error";
        public string ProcessingTimeoutTitle { get; } = "Your files takes too long to process. Please upload smaller files or smaller amount of files";
        public string ProcessingTimeoutText { get; } = "We regret to inform you that your file(s) processing took more than 3 minutes. We cannot process it at the moment. Please try again later.";
        public string InvalidFileTitle { get; } = "Invalid file, please ensure that uploading correct file";
        public string InvalidFileText { get; } = "";
        public string BadRequestTitle { get; } = "Bad Request: invalid input data.";
        public string BadRequestText { get; } = "";
        public string OtherErrorText { get; } = "Server error";
        public string SuccessfullyUploaded { get; } = "Successfully uploaded";
        public string FileSelectMessage { get; } = "Please select a file to upload";
        public string ValidateEmailMessage { get; } = "Invalid email address";
        public string WrongRegExpMessage { get; } = "Wrong regular expression";
        public string NoSearchResultsMessage { get; } = "No search results";
        public string UnlockInvalidPassword { get; } = "Invalid password";

        public string ReportTitle { get; } = "Oops! An error has occurred.";
        public string ReportPrivateLabel { get; } = "";
        public string ReportOkButton { get; } = "Close";
        public string ReportSendButton { get; } = "Send";
        public string ReportCloseButton { get; } = "Close";
        public string ReportSuccessTitle { get; } = "";
        public string ReportSuccessText { get; } = "";
    }
}