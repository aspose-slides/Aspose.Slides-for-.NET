using Aspose.Slides.Web.API.Clients.DTO.Response;
using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Web.API.Models;
using Aspose.Slides.Web.Interfaces.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.API.Controllers
{
	/// <summary>
	/// Common  API controller.
	/// </summary>
	[ApiController]
	public sealed class CommonController : ControllerBase
	{
		private const int StartingFileIndex = 1;
		private const int MaxFileIndex = 999;

		private readonly ILogger _logger;
		private readonly IFileValidatorService _fileValidatorService;
		private readonly ISourceStorage _sourceStorage;
		private readonly IProcessedStorage _processedStorage;
		private readonly ITemporaryStorage _temporaryStorage;
		private readonly IFileStreamProvider _fileStreamProvider;

		/// <summary>
		/// Constructor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="fileValidatorService">File validator instance.</param>
		/// <param name="sourceStorage">Source cloud storage instance.</param>
		/// <param name="processedStorage">Processed cloud storage instance.</param>
		/// <param name="temporaryStorage">Temporary storage instance.</param>
		/// <param name="fileStreamProvider">File stream provider instance.</param>
		public CommonController(
			ILogger<CommonController> logger,
			IFileValidatorService fileValidatorService,
			ISourceStorage sourceStorage,
			IProcessedStorage processedStorage,
			ITemporaryStorage temporaryStorage,
			IFileStreamProvider fileStreamProvider)
		{
			_logger = logger;
			_fileValidatorService = fileValidatorService;
			_sourceStorage = sourceStorage;
			_processedStorage = processedStorage;
			_temporaryStorage = temporaryStorage;
			_fileStreamProvider = fileStreamProvider;
		}

		/// <summary>
		/// Uploads files.
		/// Returns collection of upload id and filename.
		/// </summary>
		/// <returns>Files upload results.</returns>
		[HttpPost]
		[Route("api/common/UploadFiles")]
		public async Task<IEnumerable<FileSafeResult>> UploadFiles(
			[FromForm] UploadRequest request,
#pragma warning disable CS1573 // Parameter has no matching param tag in the XML comment (but other parameters do)
			CancellationToken cancellationToken = default
#pragma warning restore CS1573 // Parameter has no matching param tag in the XML comment (but other parameters do)
		)
		{
			try
			{
				using var inputFolder = _temporaryStorage.GetTemporaryFolder();
				var tasks = request.UploadFileInput.Select(
					async file =>
					{
						await using var stream = file.OpenReadStream();
						await inputFolder.SaveAsync(stream, file.FileName, cancellationToken);
					}
				);
				await Task.WhenAll(tasks);

				var idUpload = request.idUpload;

				var skipped = new List<string>();
				var filteredFiles = await FilterFolderAsync(inputFolder.ToString(), skipped, cancellationToken);
				var uploaded = await UploadFolderAsync(idUpload, filteredFiles, cancellationToken);

				_logger.LogInformation(
					"Multiple files were uploaded successfully: idUpload {@idUpload}, folder: {@folder}",
					idUpload, idUpload
				);

				cancellationToken.ThrowIfCancellationRequested();

				var result = uploaded.Select(
					filename =>
						new FileSafeResult
						{
							id = idUpload,
							FileName = Path.GetFileName(filename),
							idUpload = idUpload
						}
				).ToList();

				result.AddRange(skipped.Select(
					filename =>
						new FileSafeResult
						{
							id = idUpload,
							FileName = Path.GetFileName(filename),
							idUpload = idUpload,
							IsSuccess = false,
							idError = ErrorKeys.InvalidFile.ToString(),
							Message = "Invalid file"
						})
					);

				return result;
			}
			catch (OperationCanceledException oce)
			{
				_logger.LogInformation(oce, "UploadMultipleFiles was canceled");
				return new[]
				{
					new FileSafeResult
					{
						IsSuccess = false,
						idError = ErrorKeys.BadRequest.ToString(),
						Message = oce.Message
					}
				};
			}
		}

		/// <summary>
		/// Downloads file with specified upload id and filename.
		/// </summary>
		/// <param name="id">Download id.</param>
		/// <param name="file">File name.</param>
		/// <returns>Binary stream mime type application/octet-stream</returns>	
		[HttpGet]
		[Route("api/common/DownloadFile/{id}")]
		public async Task<IActionResult> DownloadFile(
			string id, string file,
#pragma warning disable CS1573 // Parameter has no matching param tag in the XML comment (but other parameters do)
			CancellationToken cancellationToken = default
#pragma warning restore CS1573 // Parameter has no matching param tag in the XML comment (but other parameters do)
		)
		{
			await using var stream = await _processedStorage.GetStreamAsync(id, file, cancellationToken);

			if (stream == null)
			{
				return NotFound();
			}

			var outputStream = new MemoryStream();
			await stream.CopyToAsync(outputStream, cancellationToken);
			outputStream.Seek(0, SeekOrigin.Begin);

			_logger.LogInformation("File was downloaded successfully: filename {@filename}, folder: {@folder}", file, id);

			return File(outputStream, "application/octet-stream", file);
		}

		/// <summary>
		/// Sends download url to specified email.
		/// </summary>
		/// <param name="action">Action.</param>
		/// <param name="id">Download id.</param>
		/// <param name="file">File name.</param>
		/// <param name="email">Email.</param>
		[HttpPost]
		[Route("api/common/UrlToEmail")]
		public IActionResult UrlToEmail(
			string action, string id, string file, string email,
#pragma warning disable CS1573 // Parameter has no matching param tag in the XML comment (but other parameters do)
			CancellationToken cancellationToken = default
#pragma warning restore CS1573 // Parameter has no matching param tag in the XML comment (but other parameters do)
		) => Ok(SendDownloadUrlToEmail(action, id, file, email, cancellationToken));

		/// <summary>
		/// Test fail inside controller.
		/// If code not specified - returns NotImplemented
		/// </summary>
		/// <param name="id">HTTP code</param>
		/// <param name="message">Optional message</param>
		/// <returns>Throws exception</returns>
		[HttpGet]
		[Route("api/common/TestFail")]
		public ObjectResult TestFail(int? id = null, string message = null)
		{
			int httpStatusCode = id.HasValue ? id.Value : (int)HttpStatusCode.NotImplemented;

			return StatusCode(httpStatusCode, message);
		}

		private async Task<List<string>> FilterFolderAsync(string folder, IList<string> notValidFiles, CancellationToken cancellationToken)
		{
			var filteredFiles = new List<string>();
			foreach (var file in Directory.EnumerateFiles(folder))
			{
				if (await _fileValidatorService.IsValidFileAsync(file, cancellationToken))
				{
					filteredFiles.Add(file);
				}
				else
				{
					System.IO.File.Delete(file);
					notValidFiles.Add(file);

					_logger.LogInformation($"The upload file: {Path.GetFileName(file)} was removed from server!");
				}
			}

			return filteredFiles;
		}

		private async Task<IEnumerable<string>> UploadFolderAsync(string folderName, List<string> files, CancellationToken cancellationToken)
		{
			var existingFiles = (await _sourceStorage.ListFilesAsync(folderName, cancellationToken)).ToList();

			var uploads = files
				.Select(
					path => new
					{
						Path = path,
						Filename = GetAvailableFileName(Path.GetFileName(path), existingFiles)
					}
				).ToList();

			var tasks = uploads.Select( 
				async file =>
				{
					using var stream = _fileStreamProvider.GetStream(file.Path);
					await _sourceStorage.UploadAsync(folderName, file.Filename, stream, cancellationToken);
				}
			);

			await Task.WhenAll(tasks);

			return uploads.Select(u => u.Filename);
		}

		private static string GetAvailableFileName(string filename, List<string> busy, int index = 1)
		{
			var filenameWithIndex = index == StartingFileIndex
				? filename
				: $"{Path.GetFileNameWithoutExtension(filename)}_{index}{Path.GetExtension(filename)}";

			if (!busy.Contains(filenameWithIndex) || index >= MaxFileIndex)
			{
				busy.Add(filenameWithIndex);
				return filenameWithIndex;
			}

			return GetAvailableFileName(filename, busy, index + 1);
		}

		private string SendDownloadUrlToEmail(
			string action,
			string folder,
			string file,
			string email,
			CancellationToken cancellationToken = default
		)
		{
			throw new NotImplementedException("Too many dependencies from Resources. Need to refactor first.");
			//var pathProcessor = new PathProcessor(folder, file: file);

			//var url = this.Url?.Link(
			//	"DefaultApi",
			//	new
			//	{
			//		Controller = nameof(CommonController),
			//		Action = nameof(Download),
			//		folder,
			//		file
			//	}
			//) ?? "localhostUrl";

			//var productTitle = Resources[$"{product}{action}Title"];
			//var successMessage = Resources[$"{action}SuccessMessage"];

			//var emailBody = EmailManager.PopulateBody(productTitle, url, successMessage);
			//EmailManager.SendEmail(email, Configuration.FromEmailAddress, Resources["EmailTitle"], emailBody, "");
			//return Resources["SendEmailToDownloadLink"];
		}
	}
}
