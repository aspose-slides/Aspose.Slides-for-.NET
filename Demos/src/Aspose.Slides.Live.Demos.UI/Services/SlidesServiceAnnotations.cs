using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

namespace Aspose.Slides.Live.Demos.UI.Services
{
	public partial class SlidesService
	{
		/// <summary>
		/// Removes annotations from source file, saves resulted file to out file.
		/// Returns commentaries from file.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFile">Output slides file. If value is null file not saved.</param>
		/// <returns>Commentaries.</returns>
		public string[] RemoveAnnotations(string sourceFile, string outFile)
		{
			using (var presentation = new Presentation(sourceFile))
			{
				var comments =
					presentation.CommentAuthors.Any()
					? presentation.CommentAuthors
						.SelectMany(ca => ca.Comments)
						.Select(c => $"{c.Author.Name} {c.CreatedTime}{Environment.NewLine}{c.Text}")
						.ToArray()
					: null;

				foreach (var ca in presentation.CommentAuthors)
				{
					while (ca.Comments.Any())
					{
						ca.Comments.First().Remove();
					}
				}
				presentation.CommentAuthors.Clear();

				if (outFile != null)
					presentation.Save(outFile, GetFormatFromSource(sourceFile));

				return comments;
			}
		}

		
	}
}
