using Aspose.Slides.Web.Core.Infrastructure;
using Aspose.Slides.Web.Core.Services;
using Aspose.Slides.Web.Core.Services.Storage;
using Aspose.Slides.Web.Interfaces.Services;
using Aspose.Slides.Web.Core.Services.Charts;
using Aspose.Slides.Web.Core.Services.Comparison;
using Aspose.Slides.Web.Core.Services.Conversion;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Linq;
using System.Net.Http;

namespace Aspose.Slides.Web.Core.DependencyInjection
{
	public static class SetupExtensions
	{
		public static void AddLocalFileStorage(this IServiceCollection services, string root)
		{
			services.AddScoped<ISourceStorage>(sp => new SourceFileStorage(Path.Combine(root, "source")));
			services.AddScoped<IProcessedStorage>(sp => new ProcessedFileStorage(Path.Combine(root, "processed")));
		}

		public static void AddPresentationTemplates(this IServiceCollection services, string templateStorageUrl)
		{
			services.AddSingleton<IPresentationTemplateService, PresentationTemplateService>(s => new PresentationTemplateService(
				s.GetRequiredService<ILogger<PresentationTemplateService>>(), templateStorageUrl));
		}

		public static void AddSlidesServices(this IServiceCollection services, string tmpRoot, string validFileExtensions, string ffmpegPath, string licensePath)
		{
			services.AddScoped<ITemporaryStorage, TemporaryStorage>(s => new TemporaryStorage(
				s.GetRequiredService<ILogger<TemporaryStorage>>(), tmpRoot));
			
			var validFiles = validFileExtensions.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(t => t.Trim().ToLower());

			services.AddScoped<IFileValidatorService, FileValidatorService>(s => new FileValidatorService(
				s.GetRequiredService<ILogger<FileValidatorService>>(), validFiles));
			services.AddScoped<IFileStreamProvider, FileStreamProvider>();

			services.AddSingleton<HttpClient>();

			services.AddScoped<IGifEncoder, GifEncoder>();
			services.AddSingleton<ILicenseProvider, LicenseProvider>(sp => new LicenseProvider(licensePath));
			
			services.AddScoped<IAnnotationsService, AnnotationsService>();
			services.AddScoped<ChartBuilderFactory>();
			services.AddScoped<IChartsService, ChartsService>();
			services.AddScoped<IComparisonService, ComparisonService>();
			services.AddScoped<IConversionService, ConversionService>();
			services.AddScoped<IEditorService, EditorService>();
			services.AddScoped<IImportService, ImportService>();
			services.AddScoped<IMergerService, MergerService>();
			services.AddScoped<IMetadataService, MetadataService>();
			services.AddScoped<IParseService, ParseService>();

			services.AddScoped<IProtectionService, ProtectionService>();
			services.AddScoped<IRedactionService, RedactionService>();
			services.AddScoped<ISearchService, SearchService>();
			services.AddScoped<ISignatureService, SignatureService>();
			services.AddScoped<ISplitterService, SplitterService>();

			services.AddScoped<IVideoService, VideoService>(s => new VideoService(
				s.GetRequiredService<ILogger<VideoService>>(),
				ffmpegPath,
				s.GetRequiredService<ILicenseProvider>()));
			services.AddScoped<IViewerService, ViewerService>();
			services.AddScoped<IWatermarkService, WatermarkService>();
			services.AddScoped<IMacrosService, MacrosService>();
		}
	}
}
