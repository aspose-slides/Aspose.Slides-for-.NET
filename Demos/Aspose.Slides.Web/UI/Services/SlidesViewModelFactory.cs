using System.Text;
using Aspose.Slides.Web.UI.Models;
using Aspose.Slides.Web.UI.Models.Interfaces;
using Microsoft.AspNetCore.Http;

namespace Aspose.Slides.Web.UI.Services
{
    class SlidesViewModelFactory : ISlidesViewModelFactory
    {
        private const string SchemeDelimiter = "://";
        private const string Delimiter = "/";

        public IShowroomModel CreateShowroomModel(HttpRequest request)
        {
            return new ShowroomModel();
        }

        public IBaseViewModel CreateAnnotationModel(HttpRequest request, string extension)
        {
            return new BaseViewModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "Remove Annotations",
                ProductTitleSub = "Remove all annotations from your PowerPoint presentations document"
            };
        }

        public IRedactionModel CreateSearchModel(HttpRequest request, string extension)
        {
            return new RedactionModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "Search text in your PowerPoint presentations",
                ProductTitleSub = "Search text from Microsoft PowerPoint &amp; OpenDocument presentation files via regular expression matching."
            };
        }

        public IRedactionModel CreateRedactionModel(HttpRequest request, string extension)
        {
            return new RedactionModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "Redact sensitive information",
                ProductTitleSub = "Search and replace text inside Microsoft PowerPoint &amp; OpenDocument presentation files via regular expression matching."
            };
        }

        public IBaseViewModel CreateParserModel(HttpRequest request, string extension)
        {
            return new BaseViewModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "Free online document parser",
                ProductTitleSub = "Extract Text &amp; Images Microsoft PowerPoint &amp; OpenDocument presentation files."
            };
        }

        public IBaseViewModel CreateViewerModel(HttpRequest request, string extension)
        {
            return new BaseViewModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "Free Online Presentation Viewer",
                ProductTitleSub = "Upload PPT, PPTX or ODP files to view presentation slides as images."
            };
        }

        public IRedactionModel CreateUnlockModel(HttpRequest request, string extension)
        {
            return new RedactionModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "Free Online PowerPoint and ODP file Unlocker",
                ProductTitleSub = "Unlock password protected MS PowerPoint and OpenOffice presentations."
            };
        }

        public ILockModel CreateLockModel(HttpRequest request, string extension)
        {
            return new LockModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "Free Online PowerPoint and ODP file locker",
                ProductTitleSub = "Make password protected MS PowerPoint and OpenOffice presentations online."
            };
        }

        public IMetadataModel CreateMetadataModel(HttpRequest request, string extension)
        {
            return new MetadataModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "Free Online PowerPoint metadata editor",
                ProductTitleSub = "Edit PowerPoint presentation metadata online."
            };
        }

        public IBaseViewModel CreateEditorUploaderModel(HttpRequest request, string extension)
        {
            return new BaseViewModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "PowerPoint presentations editor",
                ProductTitleSub = "Upload PPT, PPTX or ODP files to edit presentation."
            };
        }

        public IWatermarkModel CreateWatermarkModel(HttpRequest request, string extension)
        {
            return new WatermarkModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "PowerPoint presentations Watermark App",
                ProductTitleSub = "Add or remove watermark to Microsoft PowerPoint presentations files."
            };
        }

        public IConversionModel CreateConversionModel(HttpRequest request, string extension)
        {
            return new ConversionModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "PowerPoint presentations converter app",
                ProductTitleSub = "Convert your PowerPoint presentations to PDF, Open Office, HTML, XPS, Image and other formats."
            };
        }

        public IMergerViewModel CreateMergerModel(HttpRequest request, string extension)
        {
            return new MergerModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "Free online document merger for PPT, PPTX &amp; ODP Files",
                ProductTitleSub = "Merge/Combine Microsoft PowerPoint &amp; OpenDocument presentation files."
            };
        }

        public ISlideshowModel CreateSlideshowModel(HttpRequest request, string folder, string fileName)
        {
            return new SlideshowModel
            {
                APIBasePath = GetBaseUrl(request),
                FolderName = folder,
                FileName = fileName
            };
        }

        public IVideoModel CreateVideoModel(HttpRequest request, string extension)
        {
            return new VideoModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "Free online presentation to video converter",
                ProductTitleSub = "Convert your presentation to video online."
            };
        }

        public ISplitterModel CreateSplitterModel(HttpRequest request, string extension)
        {
            return new SplitterModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "PowerPoint presentations Splitter App",
                ProductTitleSub = "Split Microsoft PowerPoint presentations files."
            };
        }

        public IEditorAppModel CreateEditorAppModel(HttpRequest request, string folder, string fileName)
        {
            return new EditorAppModel
            {
                APIBasePath = GetBaseUrl(request),
                FolderName = folder,
                FileName = fileName
            };
        }

        public ISignatureModel CreateSignatureModel(HttpRequest request, string extension)
        {
            return new SignatureModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "Sign presentation online",
                ProductTitleSub = "Sign presentations with a handwritten, graphical or text signature."
            };
        }

        public IChartModel CreateChartModel(HttpRequest request, string extension)
        {
            return new ChartModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "Create charts online.",
                ProductTitleSub = "Sign presentations with a handwritten, graphical or text signature."
            };
        }

        public IComparisonModel CreateComparisonModel(HttpRequest request, string extension)
        {
            return new ComparisonModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "Online presentations comparison",
                ProductTitleSub = "Compare the text contents of two Presentations documents online"
            };
        }

        public IImportModel CreateImportModel(HttpRequest request, string extension)
        {
            return new ImportModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "Convert to PowerPoint",
                ProductTitleSub = "Import the PDF, BMP, GIF, ICO, JPEG, PNG and TIFF files to PowerPoint PPT(X) and ODP document online"
            };
        }

        public IBaseViewModel CreateRemoveMacrosModel(HttpRequest request, string extension)
        {
            return new BaseViewModel
            {
                APIBasePath = GetBaseUrl(request),
                ProductTitle = "Remove Macros from PowerPoint",
                ProductTitleSub = "Remove Macros from PowerPoint PPT and ODP document online"
            };
        }

        private static string GetBaseUrl(HttpRequest request)
        {
            var scheme = request.Scheme ?? string.Empty;
            var host = request.Host.Value ?? string.Empty;
            var pathBase = request.PathBase.Value ?? string.Empty;

            return new StringBuilder()
                .Append(scheme)
                .Append(SchemeDelimiter)
                .Append(host)
                .Append(pathBase)
                .Append(Delimiter)
                .ToString();
        }
    }
}