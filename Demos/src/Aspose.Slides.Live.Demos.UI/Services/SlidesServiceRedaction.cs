using Aspose.Slides.Live.Demos.UI.Models;
using Aspose.Slides;
using Aspose.Slides.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;

namespace Aspose.Slides.Live.Demos.UI.Services
{
	public partial class SlidesService
	{
		/// <summary>
		/// Search for specified string using regular expressions inside source file, replace string with replacement text, saves resulted file to out file.
		/// Returns found lines.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFile">Output slides file.</param>
		/// <param name="options">Redaction options.</param>
		/// <returns>Found original lines. Null if query is invalid.</returns>
		public string[] Redaction(
			string sourceFile,
			string outFile,
			RedactionOptionsModel options
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
						foreach (var para in textFrame.Paragraphs)
							foreach (var port in para.Portions)
								port.Text = MatchAndReplaceText(port.Text) ?? port.Text;
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

					foreach (var ca in presentation.CommentAuthors)
						foreach (var cm in ca.Comments)
							cm.Text = MatchAndReplaceText(cm.Text) ?? cm.Text;
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
					}

					for (int j = 0; j < presProps.CountOfCustomProperties; j++)
					{
						var propName = presProps.GetCustomPropertyName(j);
						var prop = presProps[propName];
						var value = prop.ToString();
						var updatedValue = MatchAndReplaceText(value);
						if (updatedValue != null)
							presProps.SetCustomPropertyValue(propName, updatedValue);
					}
				}

				presentation.Save(outFile, GetFormatFromSource(sourceFile));
			}

			return lines.ToArray();
		}

		
	}
}
