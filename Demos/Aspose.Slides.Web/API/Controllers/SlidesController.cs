using Aspose.Slides.Web.API.Clients.DTO.Request;
using Aspose.Slides.Web.API.Clients.DTO.Response;
using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Web.Core.Helpers;
using Aspose.Slides.Web.API.Helpers;
using Aspose.Slides.Web.API.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Slides.Web.Interfaces.Services;

namespace Aspose.Slides.Web.API.Controllers
{
	/// <summary>
	/// Slides API controller.
	/// </summary>
	[ApiController]
	[Route("api/[controller]/[action]/{id?}")]
	public sealed class SlidesController : ControllerBase
	{
		private readonly ILogger _logger;
		private readonly IChartsService _chartsService;
		private readonly IComparisonService _comparisonService;
		private readonly IConversionService _conversionService;
		private readonly IPresentationTemplateService _templateService;
		private readonly IMergerService _mergerService;
		private readonly IProcessedStorage _processedStorage;
		private readonly IFileStreamProvider _fileStreamProvider;
		private readonly ISignatureService _signatureService;
		private readonly ISourceStorage _sourceStorage;
		private readonly ITemporaryStorage _temporaryStorage;
		private readonly IVideoService _videoService;
		private readonly IMetadataService _metadataService;
		private readonly IAnnotationsService _annotationsService;
		private readonly IImportService _importerService;
		private readonly IWatermarkService _watermarkService;
		private readonly IProtectionService _protectionService;
		private readonly IRedactionService _redactionService;
		private readonly ISearchService _searchService;
		private readonly ISplitterService _splitterService;
		private readonly IParseService _parseService;
		private readonly IViewerService _viewerService;
		private readonly IEditorService _editorService;
		private readonly IMacrosService _macrosService;

		/// <summary>
		/// Default constructor.
		/// </summary>
		public SlidesController(ILogger<SlidesController> logger,
			IPresentationTemplateService templateService,
			ITemporaryStorage temporaryStorage,
			ISourceStorage sourceStorage,
			IProcessedStorage processedStorage,
			IFileStreamProvider fileStreamProvider,
			IChartsService chartsService,
			IVideoService videoService,
			ISignatureService signatureService,
			IConversionService conversionService,
			IComparisonService comparisonService,
			IMergerService mergerService,
			IMetadataService metadataService,
			IAnnotationsService annotationsService,
			IImportService importerService,
			IWatermarkService watermarkService,
			IProtectionService protectionService,
			IRedactionService redactionService,
			ISearchService searchService,
			ISplitterService splitterService,
			IParseService parseService,
			IViewerService viewerService,
			IEditorService editorService,
			IMacrosService macrosService)
		{
			_logger = logger;
			_templateService = templateService;
			_temporaryStorage = temporaryStorage;
			_sourceStorage = sourceStorage;
			_processedStorage = processedStorage;
			_fileStreamProvider = fileStreamProvider;
			_chartsService = chartsService;
			_videoService = videoService;
			_signatureService = signatureService;
			_conversionService = conversionService;
			_comparisonService = comparisonService;
			_mergerService = mergerService;
			_metadataService = metadataService;
			_annotationsService = annotationsService;
			_importerService = importerService;
			_watermarkService = watermarkService;
			_protectionService = protectionService;
			_redactionService = redactionService;
			_searchService = searchService;
			_splitterService = splitterService;
			_parseService = parseService;
			_viewerService = viewerService;
			_editorService = editorService;
			_macrosService = macrosService;
		}

		/// <summary>
		/// Removes annotations from file with specified upload id and file name, extract annotations into file.
		/// Returns details about resulted zip file.
		/// </summary>
		/// <param name="id">Upload id/</param>
		/// <param name="filename">File name.</param>
		/// <returns>Resulted file details.</returns>
		[HttpPost]
		[ActionName("RemoveAnnotations")]
		public async Task<ActionResult<FileSafeResult>> RemoveAnnotations(string id, string filename, CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			using var stream = await _sourceStorage.GetStreamAsync(id, filename, cancellationToken);

			if (stream == null)
			{
				return NotFound();
			}

			var file = await inputFolder.SaveAsync(stream, filename, cancellationToken);
			using var outputFolder = _temporaryStorage.GetTemporaryFolder();

			var commentaries = await _annotationsService.RemoveAnnotationsAsync(
				file,
				Path.Combine(outputFolder.ToString(), filename),
				cancellationToken
			);

			cancellationToken.ThrowIfCancellationRequested();
			System.IO.File.WriteAllLines(
				Path.Combine(outputFolder.ToString(), "comments.txt"),
				commentaries ?? new string[] { }
			);

			cancellationToken.ThrowIfCancellationRequested();

			var archiveName = $"{Path.GetFileNameWithoutExtension(filename)}.zip";
			using var resultStream = outputFolder.GetArchiveStream();
			await _processedStorage.UploadAsync(id, archiveName, resultStream, cancellationToken);

			_logger.LogInformation(
				"RemoveAnnotations request was processed successfully {@folder} {@filename}.", id, file
			);

			return new FileSafeResult
			{
				FileName = archiveName,
				id = id
			};
		}

		/// <summary>
		/// Search for specified string using regular expressions inside file with specified upload id and file name, saves found lines into file.
		/// Returns details about resulted file.
		/// </summary>
		/// <param name="searchRequest">Request model.</param>
		/// <returns>Resulted file details.</returns>
		[HttpPost]
		[ActionName("Search")]
		public async Task<ActionResult<FileSafeResult>> Search([FromForm] SearchRequest searchRequest, CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			var fileName = searchRequest.FileNames.First();
			using var stream = await _sourceStorage.GetStreamAsync(searchRequest.id, fileName, cancellationToken);

			if (stream == null)
			{
				return NotFound();
			}

			var file = await inputFolder.SaveAsync(stream, fileName, cancellationToken);
			using var outputFolder = _temporaryStorage.GetTemporaryFolder();

			var foundLines = await _searchService.SearchAsync(
				file,
				searchRequest.Query,
				cancellationToken
			);

			_logger.LogInformation("Search request was processed successfully {@folder} {@filename}.", searchRequest.id, fileName);

			if (foundLines == null)
			{
				return new FileSafeResult
				{
					IsSuccess = false,
					idError = ErrorKeys.InvalidReg.ToString()
				};
			}

			if (foundLines.Length < 1)
			{
				return new FileSafeResult
				{
					IsSuccess = false,
					idError = ErrorKeys.NotFound.ToString()
				};
			}

			cancellationToken.ThrowIfCancellationRequested();

			var resultFilename = "foundLines.txt";
			using var resultStream = await WriteAllToStreamAsync(foundLines);
			await _processedStorage.UploadAsync(searchRequest.id, resultFilename, resultStream, cancellationToken);

			cancellationToken.ThrowIfCancellationRequested();

			return new FileSafeResult
			{
				FileName = resultFilename,
				id = searchRequest.id
			};
		}

		/// <summary>
		/// Adds text watermark into file with specified upload id and file name.		
		/// </summary>
		/// <param name="textWatermarkOptionsRequest">Watermark options.</param>
		/// <returns>Resulted file details.</returns>
		[HttpPost]
		[ActionName("AddTextWatermark")]
		public async Task<ActionResult<IEnumerable<FileSafeResult>>> AddTextWatermark(
			[FromForm] TextWatermarkOptionsRequest textWatermarkOptionsRequest,
			CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			using var outputFolder = _temporaryStorage.GetTemporaryFolder();
			var inputFiles = new List<string>();
			var resultFiles = new List<string>();

			foreach(var sourceFile in textWatermarkOptionsRequest.FileNames)
			{
				using var stream = await _sourceStorage.GetStreamAsync(textWatermarkOptionsRequest.id, sourceFile, cancellationToken);

				if (stream == null)
				{
					return NotFound();
				}

				var inputFile = await inputFolder.SaveAsync(stream, sourceFile, cancellationToken);

				inputFiles.Add(inputFile);

				var resultFile = Path.Combine(outputFolder.ToString(), inputFile);
				resultFiles.Add(resultFile);
			}

			await _watermarkService.AddTextWatermarkAsync(
				inputFiles,
				resultFiles,
				textWatermarkOptionsRequest.GetOptions(),
				cancellationToken
			);

			var processedFiles = string.Join(",", textWatermarkOptionsRequest.FileNames);

			_logger.LogInformation("AddTextWatermark request was processed successfully {@folder} {@filename}.", textWatermarkOptionsRequest.id, processedFiles);

			cancellationToken.ThrowIfCancellationRequested();

			var result = new List<FileSafeResult>();

			foreach(var resultFile in resultFiles)
			{
				using var resultStream = _fileStreamProvider.GetStream(resultFile);				
				var fileName = textWatermarkOptionsRequest.FileNames[resultFiles.IndexOf(resultFile)];
				await _processedStorage.UploadAsync(textWatermarkOptionsRequest.id, fileName, resultStream, cancellationToken);

				result.Add(new FileSafeResult
				{
					FileName = fileName,
					id = textWatermarkOptionsRequest.id
				});
			}

			return result;
		}

		/// <summary>
		/// Adds image watermark into file with specified upload id and file name.		
		/// </summary>
		/// <param name="imageWatermarkOptionsRequest">Watermark options.</param>
		/// <returns>Resulted file details.</returns>
		[HttpPost]
		[ActionName("AddImageWatermark")]
		public async Task<ActionResult<IEnumerable<FileSafeResult>>> AddImageWatermark(
			[FromForm] ImageWatermarkOptionsRequest imageWatermarkOptionsRequest,
			CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			var imageFileName = imageWatermarkOptionsRequest.FileNames.First();
			using var imageStream = await _sourceStorage.GetStreamAsync(imageWatermarkOptionsRequest.id, imageFileName, cancellationToken);

			if (imageStream == null)
			{
				return NotFound();
			}

			imageWatermarkOptionsRequest.ImageFile = await inputFolder.SaveAsync(imageStream, imageFileName, cancellationToken);

			var inputFiles = new List<string>();
			var resultFiles = new List<string>();
			using var outputFolder = _temporaryStorage.GetTemporaryFolder();

			foreach (var sourceFile in imageWatermarkOptionsRequest.MainFileNames)
			{
				using var stream = await _sourceStorage.GetStreamAsync(imageWatermarkOptionsRequest.idMain, sourceFile, cancellationToken);

				if (stream == null)
				{
					return NotFound();
				}

				var inputFile = await inputFolder.SaveAsync(stream, sourceFile, cancellationToken);
				inputFiles.Add(inputFile);
				
				var resultFile = Path.Combine(outputFolder.ToString(), inputFile);
				resultFiles.Add(resultFile);
			}

			await _watermarkService.AddImageWatermarkAsync(
				inputFiles,
				resultFiles,
				imageWatermarkOptionsRequest.GetOptions(),
				cancellationToken
			);

			var result = new List<FileSafeResult>();

			foreach (var resultFile in resultFiles)
			{
				using var resultStream = _fileStreamProvider.GetStream(resultFile);
				var fileName = imageWatermarkOptionsRequest.MainFileNames[resultFiles.IndexOf(resultFile)];
				await _processedStorage.UploadAsync(imageWatermarkOptionsRequest.idMain, fileName, resultStream, cancellationToken);

				result.Add(new FileSafeResult
				{
					FileName = fileName,
					id = imageWatermarkOptionsRequest.idMain
				});
			}

			var processedFiles = string.Join(",", imageWatermarkOptionsRequest.MainFileNames);

			_logger.LogInformation("AddImageWatermark request was processed successfully {@folder} {@filename}.", imageWatermarkOptionsRequest.id, processedFiles);

			cancellationToken.ThrowIfCancellationRequested();

			return result;
		}

		/// <summary>
		/// Removes watermark from file with specified upload id and file name.		
		/// </summary>
		/// <param name="baseRequest">Request model.</param>
		/// <returns>Resulted file details.</returns>
		[HttpPost]
		[ActionName("RemoveWatermark")]
		public async Task<ActionResult<IEnumerable<FileSafeResult>>> RemoveWatermark(
			[FromForm] BaseRequest baseRequest,
			CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			using var outputFolder = _temporaryStorage.GetTemporaryFolder();

			var inputFiles = new List<string>();
			var resultFiles = new List<string>();

			foreach (var sourceFile in baseRequest.FileNames)
			{
				using var stream = await _sourceStorage.GetStreamAsync(baseRequest.id, sourceFile, cancellationToken);

				if (stream == null)
				{
					return NotFound();
				}

				var inputFile = await inputFolder.SaveAsync(stream, sourceFile, cancellationToken);

				inputFiles.Add(inputFile);

				var resultFile = Path.Combine(outputFolder.ToString(), inputFile);

				resultFiles.Add(resultFile);
			}

			await _watermarkService.RemoveWatermarkAsync(
				inputFiles,
				resultFiles,
				cancellationToken
			);

			var result = new List<FileSafeResult>();

			foreach (var resultFile in resultFiles)
			{
				using var resultStream = _fileStreamProvider.GetStream(resultFile);
				var fileName = baseRequest.FileNames[resultFiles.IndexOf(resultFile)];
				await _processedStorage.UploadAsync(baseRequest.id, fileName, resultStream, cancellationToken);

				result.Add(new FileSafeResult
				{
					FileName = fileName,
					id = baseRequest.id
				});
			}

			var processedFiles = string.Join(",", baseRequest.FileNames);

			_logger.LogInformation("RemoveWatermark request was processed successfully {@folder} {@filename}.", baseRequest.id, processedFiles);

			cancellationToken.ThrowIfCancellationRequested();

			return result;
		}

		/// <summary>
		/// Search for specified string using regular expressions inside file with specified upload id and file name, replaces string with replacement text, saves resulted file to out file.
		/// Returns details about resulted file.
		/// </summary>
		/// <param name="redactionOptionsRequest">Request model.</param>
		/// <returns>Resulted file details.</returns>
		[HttpPost]
		[ActionName("Redaction")]
		public async Task<ActionResult<FileSafeResult>> Redaction(
			[FromForm] RedactionOptionsRequest redactionOptionsRequest,
			CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			var fileName = redactionOptionsRequest.FileNames.First();
			using var stream = await _sourceStorage.GetStreamAsync(redactionOptionsRequest.id, fileName, cancellationToken);

			if (stream == null)
			{
				return NotFound();
			}

			var file = await inputFolder.SaveAsync(stream, fileName, cancellationToken);
			using var outputFolder = _temporaryStorage.GetTemporaryFolder();
			var resultFile = Path.Combine(outputFolder.ToString(), fileName);

			var foundLines = await _redactionService.RedactionAsync(
				file,
				resultFile,
				redactionOptionsRequest.GetOptions(),
				cancellationToken
			);

			cancellationToken.ThrowIfCancellationRequested();

			if (foundLines == null)
			{
				return new FileSafeResult
				{
					IsSuccess = false,
					idError = ErrorKeys.InvalidReg.ToString()
				};
			}

			if (foundLines.Length < 1)
			{
				return new FileSafeResult
				{
					IsSuccess = false,
					idError = ErrorKeys.NotFound.ToString()
				};
			}

			cancellationToken.ThrowIfCancellationRequested();

			using var resultStream = _fileStreamProvider.GetStream(resultFile);
			await _processedStorage.UploadAsync(redactionOptionsRequest.id, fileName, resultStream, cancellationToken);

			_logger.LogInformation("Redaction request was processed successfully {@folder} {@filename}.", redactionOptionsRequest.id, fileName);

			return new FileSafeResult
			{
				FileName = fileName,
				id = redactionOptionsRequest.id
			};
		}

		/// <summary>
		/// Parse from file with specified upload id and file name, extract text and media into separate files.
		/// Returns details about resulted zip file.
		/// </summary>
		/// <param name="id">The folder id</param>
		/// <param name="filename">The file name</param>
		/// <param name="cancellationToken"></param>
		/// <returns>Resulted file details.</returns>
		[HttpPost]
		[ActionName("Parser")]
		public async Task<ActionResult<FileSafeResult>> Parser(string id, string filename, CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			using var stream = await _sourceStorage.GetStreamAsync(id, filename, cancellationToken);

			if (stream == null)
			{
				return NotFound();
			}

			var file = await inputFolder.SaveAsync(stream, filename, cancellationToken);
			using var outputFolder = _temporaryStorage.GetTemporaryFolder();

			await _parseService.ParserAsync(
				outputFolder.ToString(),
				cancellationToken,
				file
			);

			cancellationToken.ThrowIfCancellationRequested();

			var archiveName = $"{Path.GetFileNameWithoutExtension(filename)}.zip";
			using var resultStream = outputFolder.GetArchiveStream();
			await _processedStorage.UploadAsync(id, archiveName, resultStream, cancellationToken);

			cancellationToken.ThrowIfCancellationRequested();
			_logger.LogInformation("Parser request was processed successfully {@folder} {@filename}.", id, file);

			return new FileSafeResult
			{
				FileName = archiveName,
				id = id
			};
		}

		/// <summary>
		/// Converts source file with specified upload id and file name into target format.
		/// Returns details about resulted file.
		/// </summary>
		/// <param name="conversionOptionsRequest">Request model.</param>
		/// <returns>Resulted file details.</returns>
		[HttpPost]
		[ActionName("Conversion")]
		public async Task<ActionResult<FileSafeResult>> Conversion(
			[FromForm] ConversionOptionsRequest conversionOptionsRequest,
			CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			using var outputFolder = _temporaryStorage.GetTemporaryFolder();
			var inputFiles = new List<string>();
			
			foreach (var sourceFile in conversionOptionsRequest.FileNames)
			{
				using var stream = await _sourceStorage.GetStreamAsync(conversionOptionsRequest.id, sourceFile, cancellationToken);

				if (stream == null)
				{
					return NotFound();
				}

				var inputFile = await inputFolder.SaveAsync(stream, sourceFile, cancellationToken);

				inputFiles.Add(inputFile);
			}

			var resultFiles = (await _conversionService.ConversionAsync(
				inputFiles,
				outputFolder.ToString(),
				conversionOptionsRequest.Format.ToSlidesConversionFormats(),
				cancellationToken
			)).ToList();

			cancellationToken.ThrowIfCancellationRequested();

			FileSafeResult result;

			if (resultFiles.Count > 1)
			{
				var archiveName = $"{Path.GetFileNameWithoutExtension(conversionOptionsRequest.FileNames.First())}.zip";
				using var resultStream = outputFolder.GetArchiveStream();
				await _processedStorage.UploadAsync(conversionOptionsRequest.id, archiveName, resultStream, cancellationToken);

				result = new FileSafeResult
				{
					FileName = archiveName,
					id = conversionOptionsRequest.id
				};
			}
			else
			{
				using var resultStream = _fileStreamProvider.GetStream(resultFiles.First());
				await _processedStorage.UploadAsync(conversionOptionsRequest.id, Path.GetFileName(resultFiles.First()), resultStream, cancellationToken);

				result = new FileSafeResult
				{
					FileName = Path.GetFileName(resultFiles.First()),
					id = conversionOptionsRequest.id
				};
			}

			cancellationToken.ThrowIfCancellationRequested();

			var processedFiles = string.Join(", ", conversionOptionsRequest.FileNames);
			var inputFormats = string.Join(", ", conversionOptionsRequest.FileNames.Select(f => Path.GetExtension(f)));

			_logger.LogInformation(
				"Conversion request was processed successfully {@folder} {@filename}. Input format: {@inputFormat}, output format: {@outputFormat}",
				conversionOptionsRequest.id, processedFiles, inputFormats, conversionOptionsRequest.Format
			);

			return result;
		}

		/// <summary>
		/// Merge source documents files with specified upload id into one file.
		/// Returns details about resulted file.
		/// </summary>
		/// <param name="mergerOptionsRequest">Request model.</param>
		/// <returns>Resulted file details.</returns>
		[HttpPost]
		[ActionName("Merger")]
		public async Task<ActionResult<FileSafeResult>> Merger(
			[FromForm] MergerOptionsRequest mergerOptionsRequest,
			CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			var mainFiles = new List<string>();

			foreach (var filename in mergerOptionsRequest.FileNames)
			{
				using var mainStream = await _sourceStorage.GetStreamAsync(mergerOptionsRequest.id, filename, cancellationToken);

				if (mainStream == null)
				{
					return NotFound();
				}

				// ReSharper disable once AccessToDisposedClosure
				var mainFile = await inputFolder.SaveAsync(mainStream, filename, cancellationToken);

				mainFiles.Add(mainFile);
			}

			string styleMaster = null;
			if (mergerOptionsRequest.idStyleMaster != null && mergerOptionsRequest.FileNameStyleMaster != null)
			{
				using var stream = await _sourceStorage.GetStreamAsync(
					mergerOptionsRequest.idStyleMaster, mergerOptionsRequest.FileNameStyleMaster,
					cancellationToken
				);

				if (stream == null)
				{
					return NotFound();
				}

				styleMaster = await inputFolder.SaveAsync(stream, mergerOptionsRequest.FileNameStyleMaster, cancellationToken);
			}

			using var outputFolder = _temporaryStorage.GetTemporaryFolder();

			var resultFiles = (await _mergerService.MergerAsync(
				mainFiles.ToArray(),
				outputFolder.ToString(),
				string.IsNullOrEmpty(mergerOptionsRequest.Format)
					? SlidesConversionFormats.pptx
					: mergerOptionsRequest.Format.ToSlidesConversionFormats(),
				styleMaster,
				cancellationToken
			)).ToList();

			cancellationToken.ThrowIfCancellationRequested();

			FileSafeResult result;

			if (resultFiles.Count > 1)
			{
				var archiveName = $"{Path.GetFileNameWithoutExtension(mainFiles.First())}.zip";
				using var resultStream = outputFolder.GetArchiveStream();
				await _processedStorage.UploadAsync(mergerOptionsRequest.id, archiveName, resultStream, cancellationToken);

				result = new FileSafeResult
				{
					FileName = archiveName,
					id = mergerOptionsRequest.id
				};
			}
			else
			{
				using var resultStream = _fileStreamProvider.GetStream(resultFiles.First());
				await _processedStorage.UploadAsync(
					mergerOptionsRequest.id, Path.GetFileName(resultFiles.First()), resultStream, cancellationToken
				);

				result = new FileSafeResult
				{
					FileName = Path.GetFileName(resultFiles.First()),
					id = mergerOptionsRequest.id
				};
			}

			_logger.LogInformation("Merger request was processed successfully {@folder}.", mergerOptionsRequest.id);

			return result;
		}

		/// <summary>
		/// Removes password protection from file with specified upload id and file name.
		/// Method tries to remove readonly and view protection with specified password.
		/// </summary>
		/// <param name="unProtectOptionsRequest">Request model.</param>
		/// <returns>Resulted file details.</returns>
		[HttpPost]
		[ActionName("Unlock")]
		public async Task<ActionResult<IEnumerable<FileSafeResult>>> Unlock(
			[FromForm] UnProtectOptionsRequest unProtectOptionsRequest,
			CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			using var outputFolder = _temporaryStorage.GetTemporaryFolder();
			var inputFiles = new List<string>();
			var resultFiles = new List<string>();

			foreach (var sourceFile in unProtectOptionsRequest.FileNames)
			{
				using var stream = await _sourceStorage.GetStreamAsync(unProtectOptionsRequest.id, sourceFile, cancellationToken);

				if (stream == null)
				{
					return NotFound();

				}

				var inputFile = await inputFolder.SaveAsync(stream, sourceFile, cancellationToken);

				inputFiles.Add(inputFile);

				var resultFile = Path.Combine(outputFolder.ToString(), inputFile);

				resultFiles.Add(resultFile);
			}

			await _protectionService.UnlockAsync(
				inputFiles,
				resultFiles,
				unProtectOptionsRequest.Password,
				cancellationToken
			);

			var result = new List<FileSafeResult>();

			foreach (var resultFile in resultFiles)
			{
				using var resultStream = _fileStreamProvider.GetStream(resultFile);
				var fileName = unProtectOptionsRequest.FileNames[resultFiles.IndexOf(resultFile)];
				await _processedStorage.UploadAsync(unProtectOptionsRequest.id, fileName, resultStream, cancellationToken);

				result.Add(new FileSafeResult
				{
					FileName = fileName,
					id = unProtectOptionsRequest.id
				});
			}

			cancellationToken.ThrowIfCancellationRequested();

			var processedFiles = string.Join(",", unProtectOptionsRequest.FileNames);

			_logger.LogInformation("Unlock request was processed successfully {@folder} {@filename}.", unProtectOptionsRequest.id, processedFiles);

			return result;
		}

		/// <summary>
		/// Applies protection to file with specified upload id and file name.
		/// Method adds view/edit protection with specified password and applies read-only/final flag.
		/// </summary>
		/// <param name="protectOptionsRequest">Request model.</param>
		/// <returns>Resulted file details.</returns>
		[HttpPost]
		[ActionName("Lock")]
		public async Task<ActionResult<IEnumerable<FileSafeResult>>> Lock(
			[FromForm] ProtectOptionsRequest protectOptionsRequest,
			CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			using var outputFolder = _temporaryStorage.GetTemporaryFolder();
			var inputFiles = new List<string>();
			var resultFiles = new List<string>();

			foreach (var sourceFile in protectOptionsRequest.FileNames)
			{
				using var stream = await _sourceStorage.GetStreamAsync(protectOptionsRequest.id, sourceFile, cancellationToken);

				if (stream == null)
				{
					return NotFound();
				}

				var inputFile = await inputFolder.SaveAsync(stream, sourceFile, cancellationToken);

				inputFiles.Add(inputFile);

				var resultFile = Path.Combine(outputFolder.ToString(), inputFile);

				resultFiles.Add(resultFile);
			}

			await _protectionService.LockAsync(
				inputFiles,
				resultFiles,
				protectOptionsRequest.MarkAsReadonly,
				protectOptionsRequest.MarkAsFinal,
				protectOptionsRequest.PasswordEdit,
				protectOptionsRequest.PasswordView,
				cancellationToken
			);

			var result = new List<FileSafeResult>();

			foreach (var resultFile in resultFiles)
			{
				using var resultStream = _fileStreamProvider.GetStream(resultFile);
				var fileName = protectOptionsRequest.FileNames[resultFiles.IndexOf(resultFile)];
				await _processedStorage.UploadAsync(protectOptionsRequest.id, fileName, resultStream, cancellationToken);

				result.Add(new FileSafeResult
				{
					FileName = fileName,
					id = protectOptionsRequest.id
				});
			}

			cancellationToken.ThrowIfCancellationRequested();

			var processedFiles = string.Join(",", protectOptionsRequest.FileNames);

			_logger.LogInformation("Lock request was processed successfully {@folder} {@filename}.", protectOptionsRequest.id, processedFiles);

			return result;
		}

		/// <summary>
		/// Extracts metadata from presentation with specified upload id and file name.
		/// </summary>
		/// <param name="baseRequest">Request model.</param>
		/// <returns>Presentation metadata.</returns>
		[HttpPost]
		[ActionName("GetMetadata")]
		public async Task<ActionResult<MetadataResult>> GetMetadata([FromForm] BaseRequest baseRequest, CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			var fileName = baseRequest.FileNames.First();
			var stream = await _sourceStorage.GetStreamAsync(baseRequest.id, fileName, cancellationToken);

			if (stream == null)
			{
				return NotFound();
			}

			var file = await inputFolder.SaveAsync(stream, fileName, cancellationToken);

			var metadata = _metadataService.GetMetadata(
				file,
				cancellationToken
			);

			cancellationToken.ThrowIfCancellationRequested();
			_logger.LogInformation("GetMetadata request was processed successfully {@folder} {@filename}.", baseRequest.id, fileName);

			return await Task.FromResult(
				new MetadataResult
				{
					id = baseRequest.id,
					FileName = fileName,
					Metadata = metadata.GetDTO()
				}
			);
		}

		/// <summary>
		/// Applies metadata to presentation file with the specified upload id and file name.
		/// </summary>
		/// <param name="metadataUpdateRequest">Request model.</param>
		/// <returns>Resulted file details.</returns>
		[HttpPost]
		[ActionName("UpdateMetadata")]
		public async Task<ActionResult<FileSafeResult>> UpdateMetadata(
			[FromBody] MetadataUpdateRequest metadataUpdateRequest,
			CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			var fileName = metadataUpdateRequest.FileNames.First();
			using var stream = await _sourceStorage.GetStreamAsync(metadataUpdateRequest.id, fileName, cancellationToken);

			if (stream == null)
			{
				return NotFound();
			}

			var file = await inputFolder.SaveAsync(stream, fileName, cancellationToken);
			using var outputFolder = _temporaryStorage.GetTemporaryFolder();
			var outputFile = Path.Combine(outputFolder.ToString(), Path.GetFileName(file));

			_metadataService.UpdateMetadata(
				file,
				outputFile,
				metadataUpdateRequest.Metadata.GetModel(),
				cancellationToken
			);

			cancellationToken.ThrowIfCancellationRequested();

			using var resultStream = _fileStreamProvider.GetStream(outputFile);
			await _processedStorage.UploadAsync(metadataUpdateRequest.id, Path.GetFileName(outputFile), resultStream, cancellationToken);

			_logger.LogInformation("UpdateMetadata request was processed successfully {@folder} {@filename}.", metadataUpdateRequest.id, fileName);

			return await Task.FromResult(
				new FileSafeResult
				{
					FileName = fileName,
					id = metadataUpdateRequest.id
				}
			);
		}

		/// <summary>
		/// Generates presentation HTML representation and thumbnails and returns slides information.
		/// </summary>
		/// <param name="viewerRequest">The request model.</param>
		/// <returns>The presentation information.</returns>
		[HttpPost]
		[ActionName("ViewerInfo")]
		public async Task<ActionResult<ViewerInfoResult>> GetViewerInfo([FromBody] ViewerRequest viewerRequest, CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			string file = null;
			var fileName = viewerRequest.FileNames.First();

			if (await _sourceStorage.IsExistAsync(viewerRequest.FolderName, fileName, cancellationToken))
			{
				using var stream = await _sourceStorage.GetStreamAsync(viewerRequest.FolderName, fileName, cancellationToken);

				if (stream == null)
				{
					return NotFound();
				}

				file = await inputFolder.SaveAsync(stream, fileName, cancellationToken);
			}

			using var outputFolder = _temporaryStorage.GetTemporaryFolder();
			var outputFile = Path.Combine(outputFolder.ToString(), fileName);
			var markerFile = $"{fileName}{_viewerService.MarkerId}";

			if (await _processedStorage.IsExistAsync(viewerRequest.FolderName, markerFile, cancellationToken))
			{
				using var markerStream = await _processedStorage.GetStreamAsync(
					viewerRequest.FolderName, markerFile, cancellationToken
				);

				if (markerStream == null)
				{
					return NotFound();
				}

				await outputFolder.SaveAsync(markerStream, markerFile, cancellationToken);
			}

			if (await _processedStorage.IsExistAsync(viewerRequest.FolderName, fileName, cancellationToken))
			{
				using var processedStream =
					await _processedStorage.GetStreamAsync(viewerRequest.FolderName, fileName, cancellationToken);

				if (processedStream == null)
				{
					return NotFound();
				}

				await outputFolder.SaveAsync(processedStream, fileName, cancellationToken);
			}

			var info = _viewerService.GetViewerInfo(
				viewerRequest.FolderName,
				file,
				outputFile,
				cancellationToken
			);

			await UploadFolderAsync(viewerRequest.FolderName, outputFolder.ToString(), cancellationToken);

			cancellationToken.ThrowIfCancellationRequested();

			_logger.LogInformation("ViewerInfo request was processed successfully {@folder} {@filename}.", viewerRequest.FolderName, fileName);

			if (info == null)
			{
				return await Task.FromResult(
					new ViewerInfoResult
					{
						IsSuccess = false,
						idError = ErrorKeys.InvalidFile.ToString()
					}
				);
			}

			return await Task.FromResult(new ViewerInfoResult { IsSuccess = true, Info = info.GetDTO() });
		}

		/// <summary>
		/// Returns SVG representation of the requested slide.
		/// </summary>
		[HttpGet]
		[ActionName("ViewerSlide")]
		public async Task<IActionResult> GetViewerSlide(string id, string fileName, int slide,
			CancellationToken cancellationToken = default)
		{
			var file = $"{fileName}.slide_{slide}.svg";
			await using var stream = await _processedStorage.GetStreamAsync(id, file, cancellationToken);

			if (stream == null)
			{
				return NotFound();
			}

			var outputStream = new MemoryStream();
			await stream.CopyToAsync(outputStream, cancellationToken);
			outputStream.Seek(0, SeekOrigin.Begin);

			_logger.LogInformation("ViewerSlide request was processed successfully {@folder} {@slideNumber}.", id, slide);
			return File(outputStream, "image/svg+xml", file);
		}

		/// <summary>
		/// Splits presentation to parts and saves each part to the specified format.
		/// </summary>
		/// <param name="splitterRequest">The request model</param>
		/// <returns>The resulting file archive.</returns>
		[HttpPost]
		[ActionName("Splitter")]
		public async Task<ActionResult<FileSafeResult>> Splitter([FromForm] SplitterRequest splitterRequest, CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			var fileName = splitterRequest.FileNames.First();
			using var stream = await _sourceStorage.GetStreamAsync(splitterRequest.id, fileName, cancellationToken);

			if (stream == null)
			{
				return NotFound();
			}

			var file = await inputFolder.SaveAsync(stream, fileName, cancellationToken);
			using var outputFolder = _temporaryStorage.GetTemporaryFolder();

			_splitterService.Split(
				file,
				outputFolder.ToString(),
				splitterRequest.Format.ToSlidesConversionFormats(), 
				splitterRequest.SplitType,
				splitterRequest.SplitNumber,
				splitterRequest.SplitRange,
				cancellationToken);

			var archiveName = $"{Path.GetFileNameWithoutExtension(file)}.zip";
			using var resultStream = outputFolder.GetArchiveStream();
			await _processedStorage.UploadAsync(splitterRequest.id, archiveName, resultStream, cancellationToken);

			_logger.LogInformation("Splitter request was processed successfully {@folder} {@filename}.", splitterRequest.id, fileName);

			return new FileSafeResult
			{
				FileName = archiveName,
				id = splitterRequest.id
			};
		}

		/// <summary>
		/// Splits presentation to parts and saves each part to the specified format.
		/// </summary>
		/// <param name="videoRequest">The request model</param>
		/// <returns>The resulting file archive.</returns>
		[HttpPost]
		[ActionName("Video")]
		public async Task<ActionResult<FileSafeResult>> Video([FromForm] VideoRequest videoRequest, CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			var fileName = videoRequest.FileNames.First();
			using var stream = await _sourceStorage.GetStreamAsync(videoRequest.id, fileName, cancellationToken);

			if (stream == null)
			{
				return NotFound();
			}

			var file = await inputFolder.SaveAsync(stream, fileName, cancellationToken);
			using var outputFolder = _temporaryStorage.GetTemporaryFolder();

			var resultPath = _videoService.Encode(
				file,
				outputFolder.ToString(),
				videoRequest.SplitRange,
				videoRequest.TransitionTime,
				videoRequest.VideoCodec,
				cancellationToken
			);

			using var resultStream = _fileStreamProvider.GetStream(resultPath);
			await _processedStorage.UploadAsync(videoRequest.id, Path.GetFileName(resultPath), resultStream, cancellationToken);

			_logger.LogInformation("Video request was processed successfully {@folder} {@filename}.", videoRequest.id, fileName);

			return new FileSafeResult
			{
				FileName = Path.GetFileName(resultPath),
				id = videoRequest.id
			};
		}

		/// <summary>
		/// Adds signature to the each presentation slide.
		/// </summary>
		/// <param name="signatureRequest">The request model</param>
		/// <returns>The signed presentation file or file archive.</returns>
		[HttpPost]
		[ActionName("Signature")]
		public async Task<ActionResult<FileSafeResult>> Signature([FromForm] SignatureRequest signatureRequest, CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			var fileName = signatureRequest.FileNames.First();
			using var stream = await _sourceStorage.GetStreamAsync(signatureRequest.id, fileName, cancellationToken);

			if (stream == null)
			{
				return NotFound();
			}

			var file = await inputFolder.SaveAsync(stream, fileName, cancellationToken);
			using var outputFolder = _temporaryStorage.GetTemporaryFolder();

			Stream image = null;
			try
			{
				if (!string.IsNullOrEmpty(signatureRequest.Drawing))
				{
					image = new MemoryStream(Convert.FromBase64String(signatureRequest.Drawing));
				}
				else if (!string.IsNullOrEmpty(signatureRequest.idSignatureImage))
				{
					image = await _sourceStorage.GetStreamAsync(
						signatureRequest.idSignatureImage, signatureRequest.FileNameSignatureImage, cancellationToken
					);
				}

				var resultFiles = _signatureService.Sign(
					file,
					outputFolder.ToString(),
					signatureRequest.Format.ToSlidesConversionFormats(),
					image,
					signatureRequest.Text,
					string.IsNullOrEmpty(signatureRequest.Color) ? Color.Empty : ColorTranslator.FromHtml(signatureRequest.Color),
					cancellationToken
				).ToList();

				_logger.LogInformation("Signature request was processed successfully {@folder} {@filename}.", signatureRequest.id, fileName);

				if (resultFiles.Count > 1)
				{
					var archiveName = $"{Path.GetFileNameWithoutExtension(resultFiles.First())}.zip";
					using var resultStream = outputFolder.GetArchiveStream();
					await _processedStorage.UploadAsync(signatureRequest.id, archiveName, resultStream, cancellationToken);

					return new FileSafeResult
					{
						FileName = archiveName,
						id = signatureRequest.id
					};
				}
				else
				{
					using var resultStream = _fileStreamProvider.GetStream(resultFiles.First());
					await _processedStorage.UploadAsync(
						signatureRequest.id, Path.GetFileName(resultFiles.First()), resultStream, cancellationToken
					);

					return new FileSafeResult
					{
						FileName = Path.GetFileName(resultFiles.First()),
						id = signatureRequest.id
					};
				}
			}
			finally
			{
				image?.Dispose();
			}
		}

		/// <summary>
		/// Copies presentation from the given folder in the processed file storage to a new folder.
		/// </summary>
		/// <param name="request">The request parameters.</param>
		/// <returns>Folder ID and filename of the uploaded presentation.</returns>
		[HttpPost]
		[ActionName("CopyPresentation")]
		public async Task<ActionResult<NewPresentationResponse>> CopyPresentation([FromBody] CopyPresentationRequest request,
					CancellationToken cancellationToken = default)
		{
			using var stream = await _processedStorage.GetStreamAsync(request.Folder, request.FileName, cancellationToken);

			if (stream == null)
			{
				return NotFound();
			}

			using var destination = new MemoryStream();
			await stream.CopyToAsync(destination);
			destination.Seek(0, SeekOrigin.Begin);

			var newFolder = Guid.NewGuid().ToString();
			await _processedStorage.UploadAsync(newFolder, request.FileName, destination, cancellationToken);

			return new NewPresentationResponse
			{
				Filename = request.FileName,
				Folder = newFolder
			};
		}

		/// <summary>
		/// Creates new presentation by the given template and upload it to the processed file storage.
		/// </summary>
		/// <param name="request">The request parameters.</param>
		/// <returns>Folder ID and filename of the uploaded presentation.</returns>
		[HttpPost]
		[ActionName("NewPresentation")]
		public async Task<ActionResult<NewPresentationResponse>> NewPresentation([FromBody] NewPresentationRequest request,
					CancellationToken cancellationToken = default)
		{
			using var stream = await _templateService.GetTemplateStreamAsync(request.Template, cancellationToken);

			if(stream == null)
			{
				return NotFound();
			}

			var filename = Path.GetFileName(request.Template);
			var folder = Guid.NewGuid().ToString();
			await _processedStorage.UploadAsync(folder, filename, stream, cancellationToken);
			
			return new NewPresentationResponse
			{
				Filename = filename,
				Folder = folder
			};
		}

		/// <summary>
		/// Saves edited presentation.
		/// </summary>
		/// <param name="request">The multipart request</param>
		/// <returns>The resulting file</returns>
		[HttpPost]
		[ActionName("SavePresentation")]
		public async Task<ActionResult<FileSafeResult>> SaveEditedPresentation(
			[FromForm] SavePresentationRequest request,
			CancellationToken cancellationToken = default)
		{
			var inputFolder = _temporaryStorage.GetTemporaryFolder();
			var idUpload = request.IdUpload;
			var fileName = request.FileName;

			var tasks = request.SlidesData.Select(
				async file =>
				{
					await using var stream = file.OpenReadStream();
					return await inputFolder.SaveAsync(stream, file.FileName, cancellationToken);
				}
			);
			var slides = await Task.WhenAll(tasks);

			await using var stream = await _processedStorage.GetStreamAsync(idUpload, fileName, cancellationToken);

			if (stream == null)
			{
				return NotFound();
			}

			var file = await inputFolder.SaveAsync(stream, fileName, cancellationToken);

			using var outputFolder = _temporaryStorage.GetTemporaryFolder();
			var resultFile = Path.Combine(outputFolder.ToString(), fileName);

			await _editorService.ReplaceSlidesAsync(file, slides, resultFile, cancellationToken);

			await using var resultStream = _fileStreamProvider.GetStream(resultFile);
			await _processedStorage.UploadAsync(idUpload, fileName, resultStream, cancellationToken);

			_logger.LogInformation(
				"SavePresentation request was processed successfully {@folder} {@filename}.", idUpload, fileName
			);

			return new FileSafeResult
			{
				id = idUpload,
				FileName = Path.GetFileName(resultFile)
			};
		}

		/// <summary>
		/// Creates Chart and saves to the specified format.
		/// </summary>
		/// <param name="chartRequest">The request model</param>
		/// <returns>The resulting file.</returns>
		[HttpPost]
		[ActionName("Chart")]
		public async Task<ActionResult<FileSafeResult>> Chart([FromForm] ChartRequest chartRequest, CancellationToken cancellationToken = default)
		{
			var fileName = chartRequest.FileNames.First();
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			using var outputFolder = _temporaryStorage.GetTemporaryFolder();

			List<string> resultFiles;

			if (chartRequest.IsExternalData)
			{
				resultFiles = (await _chartsService.CreateChartAsync(
						chartRequest.ChartType,
						chartRequest.JsonData,
						chartRequest.SaveFormat,
						outputFolder.ToString(),
						Path.GetFileNameWithoutExtension(fileName),
						chartRequest.IsPreview,
						cancellationToken
					)).ToList();
			}
			else
			{
				using var stream = await _sourceStorage.GetStreamAsync(chartRequest.id, fileName, cancellationToken);

				if (stream == null)
				{
					return NotFound();
				}

				var file = await inputFolder.SaveAsync(stream, fileName, cancellationToken);

				resultFiles = (await _chartsService.CreateChartAsync(
						chartRequest.ChartType,
						file,
						chartRequest.SaveFormat,
						outputFolder.ToString(),
						chartRequest.IsPreview,
						cancellationToken
					))
					.ToList();
			}

			cancellationToken.ThrowIfCancellationRequested();

			_logger.LogInformation("Chart request was processed successfully {@folder} {@filename}.", chartRequest.id, fileName);

			if (resultFiles.Count > 1)
			{
				var archiveName = $"{Path.GetFileNameWithoutExtension(resultFiles.First())}.zip";
				using var resultStream = outputFolder.GetArchiveStream();
				await _processedStorage.UploadAsync(chartRequest.id, archiveName, resultStream, cancellationToken);

				return new FileSafeResult
				{
					FileName = archiveName,
					id = chartRequest.id
				};
			}
			else
			{
				using var resultStream = _fileStreamProvider.GetStream(resultFiles.First());
				await _processedStorage.UploadAsync(
					chartRequest.id, Path.GetFileName(resultFiles.First()), resultStream, cancellationToken
				);

				return new FileSafeResult
				{
					FileName = Path.GetFileName(resultFiles.First()),
					id = chartRequest.id
				};
			}
		}

		/// <summary>
		/// Compares two presentations and returns a diff file to the specified format.
		/// </summary>
		/// <param name="comparisonRequest">The request model</param>
		/// <returns>The resulting file.</returns>
		[HttpPost]
		[ActionName("Comparison")]
		public async Task<ActionResult<FileSafeResult>> Comparison([FromForm] ComparisonRequest comparisonRequest, CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			using var stream = await _sourceStorage.GetStreamAsync(comparisonRequest.FirstFolderId, comparisonRequest.FirstFileName, cancellationToken);

			if (stream == null)
			{
				return NotFound();
			}

			var file = await inputFolder.SaveAsync(stream, comparisonRequest.FirstFileName, cancellationToken);
			using var outputFolder = _temporaryStorage.GetTemporaryFolder();
			var outputFile = Path.Combine(outputFolder.ToString(), Path.GetFileName(file));

			using var inputFolder2 = _temporaryStorage.GetTemporaryFolder();
			using var stream2 = await _sourceStorage.GetStreamAsync(
				comparisonRequest.SecondFolderId, comparisonRequest.SecondFileName, cancellationToken
			);

			if (stream2 == null)
			{
				return NotFound();
			}

			var file2 = await inputFolder2.SaveAsync(stream2, comparisonRequest.SecondFileName, cancellationToken);

			cancellationToken.ThrowIfCancellationRequested();

			var result = await _comparisonService.ComparePresentationsAsync(
				file,
				file2,
				outputFolder.ToString(),
				comparisonRequest.SaveFormat,
				comparisonRequest.ComparisonMethod,
				cancellationToken
			);

			cancellationToken.ThrowIfCancellationRequested();

			_logger.LogInformation("Comparison request was processed successfully {@folder} {@filename}.", comparisonRequest.id, comparisonRequest.FirstFileName);

			if (result == null)
			{
				var archiveName = $"{Path.GetFileNameWithoutExtension(outputFile)}.zip";
				using var resultStream = outputFolder.GetArchiveStream();
				await _processedStorage.UploadAsync(comparisonRequest.FirstFolderId, archiveName, resultStream, cancellationToken);

				return new FileSafeResult
				{
					FileName = archiveName,
					id = comparisonRequest.FirstFolderId
				};
			}
			else
			{
				using var resultStream = _fileStreamProvider.GetStream(result);
				await _processedStorage.UploadAsync(comparisonRequest.FirstFolderId, Path.GetFileName(result), resultStream, cancellationToken);

				return new FileSafeResult
				{
					FileName = Path.GetFileName(result),
					id = comparisonRequest.FirstFolderId
				};
			}
		}

		private static async Task<MemoryStream> WriteAllToStreamAsync(string[] foundLines)
		{
			var resultStream = new MemoryStream();
			var writer = new StreamWriter(resultStream);
			Array.ForEach(foundLines, writer.WriteLine);
			await writer.FlushAsync();
			resultStream.Seek(0, SeekOrigin.Begin);

			return resultStream;
		}

		private Task UploadFolderAsync(string folderName, string folderPath, CancellationToken cancellationToken)
		{
			var tasks = Directory.EnumerateFiles(folderPath).Select(
				async file =>
				{
					using var stream = _fileStreamProvider.GetStream(file);
					await _processedStorage.UploadAsync(folderName, Path.GetFileName(file), stream, cancellationToken);
				}
			);

			return Task.WhenAll(tasks);
		}

		/// <summary>
		/// Imports files into a presentation to the specified format.
		/// </summary>
		/// <param name="importRequest">The request model</param>
		/// <returns>The resulting file.</returns>
		[HttpPost]
		[ActionName("Import")]
		public async Task<ActionResult<FileSafeResult>> Import([FromForm] ImportRequest importRequest, CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			var inputFiles = new List<string>();

			foreach (var sourceFile in importRequest.FileNames)
			{
				using var stream = await _sourceStorage.GetStreamAsync(importRequest.id, sourceFile, cancellationToken);

				if (stream == null)
				{
					return NotFound();
				}

				var inputFile = await inputFolder.SaveAsync(stream, sourceFile, cancellationToken);

				inputFiles.Add(inputFile);
			}

			cancellationToken.ThrowIfCancellationRequested();

			using var outputFolder = _temporaryStorage.GetTemporaryFolder();
			var result = await _importerService.ImportToPresentationAsync(inputFiles,
				outputFolder.ToString(),
				importRequest.SaveFormat,
				cancellationToken);

			cancellationToken.ThrowIfCancellationRequested();

			_logger.LogInformation("Import request was processed successfully {@folder} {@filename}.", importRequest.id, string.Join(", ", importRequest.FileNames));

			using var resultStream = _fileStreamProvider.GetStream(result);
			await _processedStorage.UploadAsync(importRequest.id, Path.GetFileName(result), resultStream, cancellationToken);

			return new FileSafeResult
			{
				FileName = Path.GetFileName(result),
				id = importRequest.id
			};
		}

		/// <summary>
		/// Remove all active content(Macros) from presentations
		/// </summary>
		/// <param name="baseRequest">Request model.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns>Resulted file details.</returns>
		[HttpPost]
		[ActionName("RemoveMacros")]
		public async Task<ActionResult<IEnumerable<FileSafeResult>>> RemoveMacros(
			[FromForm] BaseRequest baseRequest,
			CancellationToken cancellationToken = default)
		{
			using var inputFolder = _temporaryStorage.GetTemporaryFolder();
			using var outputFolder = _temporaryStorage.GetTemporaryFolder();
			var inputFiles = new List<string>();
			
			foreach (var sourceFile in baseRequest.FileNames)
			{
				using var stream = await _sourceStorage.GetStreamAsync(baseRequest.id, sourceFile, cancellationToken);

				if (stream == null)
				{
					return NotFound();
				}

				cancellationToken.ThrowIfCancellationRequested();

				var inputFile = await inputFolder.SaveAsync(stream, sourceFile, cancellationToken);

				inputFiles.Add(inputFile);
			}

			var resultFiles = _macrosService.RemoveMacros(
				inputFiles,
				outputFolder.ToString(),
				cancellationToken
			);

			cancellationToken.ThrowIfCancellationRequested();

			var result = new List<FileSafeResult>();

			foreach (var resultFile in resultFiles)
			{
				using var resultStream = _fileStreamProvider.GetStream(resultFile);

				cancellationToken.ThrowIfCancellationRequested();

				await _processedStorage.UploadAsync(baseRequest.id, Path.GetFileName(resultFile), resultStream, cancellationToken);

				result.Add(new FileSafeResult
				{
					FileName = Path.GetFileName(resultFile),
					id = baseRequest.id
				});
			}

			cancellationToken.ThrowIfCancellationRequested();

			var processedFiles = string.Join(",", baseRequest.FileNames);

			_logger.LogInformation("RemoveMacros request was processed successfully {@folder} {@filename}.", baseRequest.id, processedFiles);

			return result;
		}
	}
}
