using Aspose.Slides.Util;
using Aspose.Slides.Web.Core.Enums;
using Aspose.Slides.Web.Core.Helpers;
using Aspose.Slides.Web.Core.Infrastructure;
using Aspose.Slides.Web.Interfaces.Models.Redaction;
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
	/// Implementation of slides redaction logic.
	/// </summary>
	internal sealed class RedactionService : SlidesServiceBase, IRedactionService
	{
		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="licenseProvider"></param>
		public RedactionService(ILogger<RedactionService> logger, ILicenseProvider licenseProvider) : base(logger)
		{
			licenseProvider.SetAsposeLicense(AsposeProducts.Slides);
		}

		/// <summary>
		/// Search for specified string using regular expressions inside source file, replace string with replacement text, saves resulted file to out file.
		/// Returns found lines.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFile">Output slides file.</param>
		/// <param name="options">Redaction options.</param>
		/// <returns>Found original lines. Null if query is invalid.</returns>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		public string[] Redaction(
			string sourceFile,
			string outFile,
			RedactionOptions options,
			CancellationToken cancellationToken = default
		)
		{
			try
			{
				Regex.IsMatch("", options.SearchQuery);
			}
			catch (Exception)
			{
				return null;
			}

			var lines = new List<string>();

			using (var presentation = new Presentation(sourceFile))
			{
				cancellationToken.ThrowIfCancellationRequested();

				string MatchAndReplaceText(string text)
				{
					if (
						Regex.IsMatch(
							text,
							options.SearchQuery,
							options.IsCaseSensitiveSearch
								? RegexOptions.None
								: RegexOptions.IgnoreCase
						)
					)
					{
						lines.Add(text);

						return Regex.Replace(
							text,
							options.SearchQuery,
							options.ReplaceText ?? "",
							options.IsCaseSensitiveSearch
								? RegexOptions.None
								: RegexOptions.IgnoreCase
						) ?? "";
					}
					else
						return null;
				}

				if (options.MustReplaceText)
				{
					var textFrames = SlideUtil.GetAllTextFrames(presentation, true);

					foreach (var textFrame in textFrames)
					{
						foreach (var para in textFrame.Paragraphs)
						{
							foreach (var port in para.Portions)
							{
								port.Text = MatchAndReplaceText(port.Text) ?? port.Text;

								cancellationToken.ThrowIfCancellationRequested();
							}

							cancellationToken.ThrowIfCancellationRequested();
						}

						cancellationToken.ThrowIfCancellationRequested();
					}					
				}

				if (options.MustReplaceComments)
				{
					var comments =
						presentation.CommentAuthors.Any()
						? presentation.CommentAuthors
							.SelectMany(ca => ca.Comments)
							.Select(c => $"{c.Author.Name} {c.CreatedTime}{Environment.NewLine}{c.Text}")
							.ToArray()
						: null;

					cancellationToken.ThrowIfCancellationRequested();
					foreach (var ca in presentation.CommentAuthors)
					{
						foreach (var cm in ca.Comments)
						{
							cm.Text = MatchAndReplaceText(cm.Text) ?? cm.Text;

							cancellationToken.ThrowIfCancellationRequested();
						}

						cancellationToken.ThrowIfCancellationRequested();
					}
				}

				if (options.MustReplaceMetadata)
				{
					var presProps = presentation.DocumentProperties;
					var props = presProps;
					foreach (var prop in props.GetType().GetProperties())
					{
						if (prop.CanWrite && prop.PropertyType == typeof(string))
						{
							var value = prop.GetValue(props).ToString();

							var updatedValue = MatchAndReplaceText(value);
							if (updatedValue != null)
								prop.SetValue(props, updatedValue);
						}

						cancellationToken.ThrowIfCancellationRequested();
					}

					for (int j = 0; j < presProps.CountOfCustomProperties; j++)
					{
						var propName = presProps.GetCustomPropertyName(j);
						var prop = presProps[propName];
						var value = prop.ToString();
						var updatedValue = MatchAndReplaceText(value);
						if (updatedValue != null)
							presProps.SetCustomPropertyValue(propName, updatedValue);

						cancellationToken.ThrowIfCancellationRequested();
					}
				}

				cancellationToken.ThrowIfCancellationRequested();
				presentation.Save(outFile, sourceFile.GetSlidesExportSaveFormatBySourceFile());
			}

			return lines.ToArray();
		}

		/// <summary>
		/// Asynchronously search for specified string using regular expressions inside source file, replace string with replacement text, saves resulted file to out file.
		/// Returns found lines.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFile">Output slides file.</param>
		/// <param name="options">Redaction options.</param>
		/// <returns>Found original lines. Null if query is invalid.</returns>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		public async Task<string[]> RedactionAsync(
			string sourceFile,
			string outFile,
			RedactionOptions options,
			CancellationToken cancellationToken = default
		) => await Task.Run(() => Redaction(sourceFile, outFile, options, cancellationToken));
	}
}
