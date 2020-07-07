using Aspose.Slides;
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
		/// Search for specified string using regular expressions inside source file, saves found result file to out file.
		/// Returns found lines.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="query">Search query string.</param>		
		/// <returns>Found lines. Null if query is invalid.</returns>
		public string[] Search(string sourceFile, string query)
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
				foreach (var slide in presentation.Slides)
				{
					foreach (var shp in slide.Shapes)
					{
						if (shp is AutoShape ashp)
							lines.Add(ashp.TextFrame.Text);
					}

					var notes = slide.NotesSlideManager.NotesSlide?.NotesTextFrame?.Text;

					if (!string.IsNullOrEmpty(notes))
						lines.Add(notes);
				}
			}

			var resultLines = lines.Where(l => Regex.IsMatch(l, query)).ToArray();

			return resultLines;
		}

		
	}
}
