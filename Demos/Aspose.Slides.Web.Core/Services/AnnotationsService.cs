using Aspose.Slides.Web.Core.Helpers;
using Aspose.Slides.Web.Interfaces.Services;
using Aspose.Slides.Web.Core.Infrastructure;
using Aspose.Slides.Web.Core.Enums;
using Microsoft.Extensions.Logging;
using System;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Core.Services
{
	/// <summary>
	/// The implementation of Annotations service logic.
	/// </summary>
	public sealed class AnnotationsService : SlidesServiceBase, IAnnotationsService
	{
		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="licenseProvider"></param>
		public AnnotationsService(ILogger<AnnotationsService> logger, ILicenseProvider licenseProvider) : base(logger)
		{
			licenseProvider.SetAsposeLicense(AsposeProducts.Slides);
		}

		/// <summary>
		/// Removes annotations from source file, saves resulted file to out file.
		/// Returns commentaries from file.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFile">Output slides file. If value is null file not saved.</param>
		/// <returns>Commentaries.</returns>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		public string[] RemoveAnnotations(string sourceFile, string outFile, CancellationToken cancellationToken = default)
		{
			using (var presentation = new Presentation(sourceFile))
			{
				cancellationToken.ThrowIfCancellationRequested();

				var comments =
					presentation.CommentAuthors.Any()
					? presentation.CommentAuthors
						.SelectMany(ca => ca.Comments)
						.Select(c => $"{c.Author.Name} {c.CreatedTime.ToString(CultureInfo.InvariantCulture)}{Environment.NewLine}{c.Text}")
						.ToArray()
					: null;

				foreach (var ca in presentation.CommentAuthors)
				{
					while (ca.Comments.Any())
					{
						ca.Comments.First().Remove();

						cancellationToken.ThrowIfCancellationRequested();
					}

					cancellationToken.ThrowIfCancellationRequested();
				}
				presentation.CommentAuthors.Clear();

				cancellationToken.ThrowIfCancellationRequested();
				if (outFile != null)
					presentation.Save(outFile, sourceFile.GetSlidesExportSaveFormatBySourceFile());

				return comments;
			}
		}

		/// <summary>
		/// Asynchronously removes annotations from source file, saves resulted file to out file.
		/// Returns commentaries from file.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFile">Output slides file. If value is null file not saved.</param>
		/// <returns>Commentaries.</returns>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		public async Task<string[]> RemoveAnnotationsAsync(string sourceFile, string outFile, CancellationToken cancellationToken = default)
			=> await Task.Run(() => RemoveAnnotations(sourceFile, outFile, cancellationToken));
	}
}
