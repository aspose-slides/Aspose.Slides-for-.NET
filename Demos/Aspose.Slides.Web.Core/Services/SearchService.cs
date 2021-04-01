using Aspose.Slides.Web.Core.Enums;
using Aspose.Slides.Web.Core.Infrastructure;
using Aspose.Slides.Web.Interfaces.Services;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Core.Services
{
	/// <summary>
	/// Implementation of search logic.
	/// </summary>
	internal sealed class SearchService : SlidesServiceBase, ISearchService
	{
		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="licenseProvider"></param>
		public SearchService(ILogger<SearchService> logger, ILicenseProvider licenseProvider) : base(logger)
		{
			licenseProvider.SetAsposeLicense(AsposeProducts.Slides);
		}

		/// <summary>
		/// Search for specified string using regular expressions inside source file, saves found result file to out file.
		/// Returns found lines.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="query">Search query string.</param>		
		/// <returns>Found lines. Null if query is invalid.</returns>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		public string[] Search(string sourceFile, string query, CancellationToken cancellationToken = default)
		{
			try
			{
				Regex.IsMatch("", query);
			}
			catch (Exception)
			{
				return null;
			}

			var lines = new List<string>();

			using (var presentation = new Presentation(sourceFile))
			{
				cancellationToken.ThrowIfCancellationRequested();
				foreach (var slide in presentation.Slides)
				{
					foreach (var shp in slide.Shapes)
					{
						if (shp is AutoShape ashp)
							lines.Add(ashp.TextFrame.Text);

						cancellationToken.ThrowIfCancellationRequested();
					}

					var notes = slide.NotesSlideManager.NotesSlide?.NotesTextFrame?.Text;

					if (!string.IsNullOrEmpty(notes))
						lines.Add(notes);

					cancellationToken.ThrowIfCancellationRequested();
				}
			}

			var resultLines = lines.Where(l => Regex.IsMatch(l, query)).ToArray();

			cancellationToken.ThrowIfCancellationRequested();
			return resultLines;
		}

		/// <summary>
		/// Asynchronously search for specified string using regular expressions inside source file, saves found result file to out file.
		/// Returns found lines.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="query">Search query string.</param>		
		/// <returns>Found lines. Null if query is invalid.</returns>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		public async Task<string[]> SearchAsync(string sourceFile, string query, CancellationToken cancellationToken = default)
			=> await Task.Run(() => Search(sourceFile, query, cancellationToken));
	}
}
