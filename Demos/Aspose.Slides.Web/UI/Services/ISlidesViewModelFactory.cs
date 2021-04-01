using Aspose.Slides.Web.UI.Models.Interfaces;
using Microsoft.AspNetCore.Http;

namespace Aspose.Slides.Web.UI.Services
{
	public interface ISlidesViewModelFactory
	{
		IShowroomModel CreateShowroomModel(HttpRequest request);
		IBaseViewModel CreateAnnotationModel(HttpRequest request, string extension);
		IRedactionModel CreateSearchModel(HttpRequest request, string extension);
		IRedactionModel CreateRedactionModel(HttpRequest request, string extension);
		IBaseViewModel CreateParserModel(HttpRequest request, string extension);
		IBaseViewModel CreateViewerModel(HttpRequest request, string extension);
		IRedactionModel CreateUnlockModel(HttpRequest request, string extension);
		ILockModel CreateLockModel(HttpRequest request, string extension);
		IMetadataModel CreateMetadataModel(HttpRequest request, string extension);
		IBaseViewModel CreateEditorUploaderModel(HttpRequest request, string extension);
		IWatermarkModel CreateWatermarkModel(HttpRequest request, string extension);
		IConversionModel CreateConversionModel(HttpRequest request, string extension);
		IMergerViewModel CreateMergerModel(HttpRequest request, string extension);
		ISlideshowModel CreateSlideshowModel(HttpRequest request, string folder, string fileName);
		IVideoModel CreateVideoModel(HttpRequest request, string extension);
		ISplitterModel CreateSplitterModel(HttpRequest request, string extension);
		IEditorAppModel CreateEditorAppModel(HttpRequest request, string folder, string fileName);
		ISignatureModel CreateSignatureModel(HttpRequest request, string extension);
		IChartModel CreateChartModel(HttpRequest request, string extension);
		IComparisonModel CreateComparisonModel(HttpRequest request, string extension);
		IImportModel CreateImportModel(HttpRequest request, string extension);
		IBaseViewModel CreateRemoveMacrosModel(HttpRequest request, string extension);
	}
}
