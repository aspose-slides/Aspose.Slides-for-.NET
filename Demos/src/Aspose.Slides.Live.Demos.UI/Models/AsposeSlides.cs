using Aspose.Slides.Live.Demos.UI.Helpers;

using Aspose.Slides.Live.Demos.UI.Services;
using System.Threading;
using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using Aspose.Slides.Live.Demos.UI.Models.Slides;

namespace Aspose.Slides.Live.Demos.UI.Models
{
	/// <summary>
	/// AsposeSlides.
	/// </summary>	
	public sealed class AsposeSlides : SlidesBase
	{
		

		private static async Task<T> ExecuteWithHandling<T>(Func<SlidesService, Task<T>> func) where T : BaseResult, new()
		{
			try
			{
				using (var sc = new SlidesService())
					return await func(sc);
			}
			catch (Exception e) when (e is PptReadException || e is ArgumentException)
			{
				return new T
				{
					IsSuccess = false,
					idError = "InvalidFile"
				};
			}
		}

		/// <summary>
		/// Removes annotations from file with specified upload id and file name, extract annotations into file.
		/// Returns details about resulted zip file.
		/// </summary>
		/// <param name="id">Upload id/</param>
		/// <param name="file">File name.</param>
		/// <returns>Resulted file details.</returns>

		public FileSafeResult RemoveAnnotations(string id, string file)
		{

			var pathProcessor = new PathProcessor(id, file: file, checkDefaultSourceFileExistence: true);
			var annotationsFile = pathProcessor.GetOutFilePath("comments.txt");
			SlidesService slidesService = new SlidesService();
			var commentaries = slidesService.RemoveAnnotations(
					pathProcessor.DefaultSourceFile,
					pathProcessor.DefaultOutFile
				);

			//System.IO.File.WriteAllLines(annotationsFile, commentaries );

			return pathProcessor.GetResultZipped();
		}


		/// <summary>
		/// Search for specified string using regular expressions inside file with specified upload id and file name, saves found lines into file.
		/// Returns details about resulted file.
		/// </summary>
		/// <param name="model">Request model.</param>		
		/// <returns>Resulted file details.</returns>		
		public FileSafeResult Search([FromBody]SearchRequestModel model) 
				{
			SlidesService slidesService = new SlidesService();
			var pathProcessor = new PathProcessor(model.id, file: model.FileName, checkDefaultSourceFileExistence: true);

					var foundLines =  slidesService.Search(
						pathProcessor.DefaultSourceFile,
						model.Query
					);

					if (foundLines == null)
						return new FileSafeResult
						{
							IsSuccess = false,
							idError = "InvalidReg"
						};

					if (foundLines.Length < 1)
						return new FileSafeResult
						{
							IsSuccess = false,
							idError = "NotFound"
						};

					var foundLinesFile = pathProcessor.GetOutFilePath("foundLines.txt");
					System.IO.File.WriteAllLines(foundLinesFile, foundLines);
					return pathProcessor.GetResult("foundLines.txt");
				}
			

		/// <summary>
		/// Adds text watermark into file with specified upload id and file name.		
		/// </summary>		
		/// <param name="model">Watermark options.</param>
		/// <returns>Resulted file details.</returns>
		
		public FileSafeResult AddTextWatermark([FromBody]TextWatermarkOptionsModel model)
		
				{
			SlidesService slidesService = new SlidesService();
			var pathProcessor = new PathProcessor(model.id, model.FileName, true);

			slidesService.AddTextWatermark(
						pathProcessor.DefaultSourceFile,
						pathProcessor.DefaultOutFile,
						model
					);

					return pathProcessor.GetResult();
				}
			

		/// <summary>		
		/// Adds image watermark into file with specified upload id and file name.		
		/// </summary>		
		/// <param name="model">Watermark options.</param>
		/// <returns>Resulted file details.</returns>
		
		public FileSafeResult AddImageWatermark([FromBody]ImageWatermarkOptionsModel model) 
			
				{
			SlidesService slidesService = new SlidesService();
			var pathProcessorImage = new PathProcessor(model.id, model.FileName, true);
					model.ImageFile = pathProcessorImage.DefaultSourceFile;

					var pathProcessor = new PathProcessor(model.idMain, model.MainFileName, true);

			slidesService.AddImageWatermark(
						pathProcessor.DefaultSourceFile,
						pathProcessor.DefaultOutFile,
						model
					);

					return pathProcessor.GetResult();
				}
			

		/// <summary>
		/// Removes watermark from file with specified upload id and file name.		
		/// </summary>
		/// <param name="model">Request model.</param>
		/// <returns>Resulted file details.</returns>

		public FileSafeResult RemoveWatermark([FromBody]BaseRequestModel model) 
				{
			SlidesService slidesService = new SlidesService();
			var pathProcessor = new PathProcessor(model.id, model.FileName, true);

			slidesService.RemoveWatermark(
						pathProcessor.DefaultSourceFile,
						pathProcessor.DefaultOutFile
					);

					return pathProcessor.GetResult();
				}
		

		/// <summary>
		/// Search for specified string using regular expressions inside file with specified upload id and file name, replaces string with replacement text, saves resulted file to out file.
		/// Returns details about resulted file.
		/// </summary>
		/// <param name="model">Request model.</param>
		/// <returns>Resulted file details.</returns>

		public FileSafeResult Redaction([FromBody]RedactionOptionsModel model)
			
				{
					var pathProcessor = new PathProcessor(model.id, model.FileName, true);
			SlidesService slidesService = new SlidesService();
			var foundLines = slidesService.Redaction(
						pathProcessor.DefaultSourceFile,
						pathProcessor.DefaultOutFile,
						model
					);

					if (foundLines == null)
						return new FileSafeResult
						{
							IsSuccess = false,
							idError = "InvalidReg"
						};

					if (foundLines.Length < 1)
						return new FileSafeResult
						{
							IsSuccess = false,
							idError = "NotFound"
						};

					return pathProcessor.GetResult();
				}
			

		/// <summary>
		/// Parse from file with specified upload id and file name, extract text and media into separate files.
		/// If file name not specified, parse all files in folder.
		/// Returns details about resulted zip file.
		/// </summary>
		/// <param name="id">Upload id/</param>
		/// <param name="file">File name.</param>
		/// <returns>Resulted file details.</returns>
	
		public FileSafeResult Parser(string id, string file) 
				{
					var pathProcessor = new PathProcessor(id, file, file != null);
			SlidesService slidesService = new SlidesService();
			slidesService.Parser(
						pathProcessor.OutFolder,
						file != null
							? pathProcessor.DefaultSourceFile
							: pathProcessor.SourceFolder
					);
					return pathProcessor.GetResultZipped();
				}
			

		/// <summary>
		/// Converts source file with specified upload id and file name into target format.		
		/// Returns details about resulted file.
		/// </summary>
		/// <param name="model">Request model.</param>
		/// <returns>Resulted file details.</returns>	
		public FileSafeResult Conversion([FromBody]ConversionOptions model)
		{
			SlidesService slidesService = new SlidesService();
			
			   var pathProcessor = new PathProcessor(model.id, model.FileName, model.FileName != null);
			   var result = slidesService.ConvertFile(
				   pathProcessor.DefaultSourceFile,
				   pathProcessor.OutFolder,
				   model.Format.ParseEnum<SlidesConversionFormat>()
			   );

			   if (result == null)
				   return pathProcessor.GetResultZipped();
			   else
				   return pathProcessor.GetResult(Path.GetFileName(result));
		   
		}

		/// <summary>
		/// Merge source documents files with specified upload id into one file.		
		/// Returns details about resulted file.
		/// </summary>
		/// <param name="model">Request model.</param>
		/// <returns>Resulted file details.</returns>
		
		public FileSafeResult Merger([FromBody]MergerOptions model) 
				{
					SlidesService slidesService = new SlidesService();
					var pathProcessor = new PathProcessor(model.idMain, null, false);

					PathProcessor pathProcessorStyleMaster = null;
					if (model.idStyleMaster != null && model.FileNameStyleMaster != null)
						pathProcessorStyleMaster = new PathProcessor(model.idStyleMaster, model.FileNameStyleMaster, true);

					var result = slidesService.Merger(
						pathProcessor.OutFolder,
						string.IsNullOrEmpty(model.Format)
							? SlidesConversionFormat.pptx
							: model.Format.ParseEnum<SlidesConversionFormat>(),
						pathProcessorStyleMaster?.DefaultSourceFile,
						model.MainFiles.Select(fileName => Path.Combine(pathProcessor.SourceFolder, fileName)).ToArray()
					);

					return pathProcessor.GetResult(Path.GetFileName(result));
				}
		


		/// <summary>
		/// Removes password protection from file with specified upload id and file name.
		/// Method tries to remove readonly and view protection with specified password.
		/// </summary>
		/// <param name="model">Request model.</param>
		/// <returns>Resulted file details.</returns>

		public FileSafeResult Unlock([FromBody]UnProtectOptions model)
		{
			SlidesService slidesService = new SlidesService();
			var pathProcessor = new PathProcessor(model.id, model.FileName, true);

					try
					{
				slidesService.UnlockFile(
							pathProcessor.DefaultSourceFile,
							pathProcessor.DefaultOutFile,
							model.Password
						);

						return pathProcessor.GetResult();
					}
					catch (InvalidPasswordException)
					{
						return new FileSafeResult
						{
							IsSuccess = false,
							idError = "InvalidPassword"
						};
					}
				
		}

		/// <summary>
		/// Applies protection to file with specified upload id and file name.
		/// Method adds view/edit protection with specified password and applies read-only/final flag.
		/// </summary>
		/// <param name="model">Request model.</param>
		/// <returns>Resulted file details.</returns>

		public FileSafeResult Lock([FromBody]ProtectOptions model)
		{
			SlidesService slidesService = new SlidesService();
			var pathProcessor = new PathProcessor(model.id, model.FileName, true);

			slidesService.LockFile(
						sourceFile: pathProcessor.DefaultSourceFile,
						outFile: pathProcessor.DefaultOutFile,
						markAsReadonly: model.MarkAsReadonly,
						markAsFinal: model.MarkAsFinal,
						passwordEdit: model.PasswordEdit,
						passwordView: model.PasswordView
					);

					return pathProcessor.GetResult();
				
		}

	
		/// <summary>
		/// Splits presentation to parts and saves each part to the specified format.
		/// </summary>
		/// <param name="model">The request model</param>
		/// <returns>The resulting file archive.</returns>
	
		public FileSafeResult Splitter([FromBody]SplitterRequestModel model,

			CancellationToken cancellationToken = default(CancellationToken)

			) 
				{
					var pathProcessor = new PathProcessor(model.id, model.FileName, true);
			SlidesService slidesService = new SlidesService();
			slidesService.Split(pathProcessor.DefaultSourceFile,
						Path.GetDirectoryName(pathProcessor.DefaultOutFile),
						model.Format.ParseEnum<SlidesConversionFormat>(),
						model.SplitType,
						model.SplitNumber, model.SplitRange, cancellationToken);

					return pathProcessor.GetResultZipped();
				}
	}
}
