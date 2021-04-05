using Aspose.Slides.Web.Interfaces.Services;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Slides.Web.Core.Infrastructure;
using Aspose.Slides.Web.Core.Enums;

namespace Aspose.Slides.Web.Core.Services
{
	/// <summary>
	/// Implementation of edit logic.
	/// </summary>
	internal sealed class EditorService : SlidesServiceBase, IEditorService
	{
		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="licenseProvider"></param>
		public EditorService(ILogger<EditorService> logger, ILicenseProvider licenseProvider) : base(logger)
		{
			licenseProvider.SetAsposeLicense(AsposeProducts.Slides);
		}

		/// <summary>
		/// Replaces slides with given svg-files in the presentation.
		/// </summary>
		/// <param name="sourceFile">The source presentation file.</param>
		/// <param name="slides">The list of given svg-files.</param>
		/// <param name="outFile">The path to the resulting file</param>
		public Task ReplaceSlidesAsync(string sourceFile, IEnumerable<string> slides, string outFile,
#pragma warning disable CS1573 // Parameter has no matching param tag in the XML comment (but other parameters do)
			CancellationToken cancellationToken = default
#pragma warning restore CS1573 // Parameter has no matching param tag in the XML comment (but other parameters do)
		)
		{
			using (var presentation = new Presentation(sourceFile))
			{
				foreach (var slide in slides)
				{
					var index = int.Parse(Path.GetFileNameWithoutExtension(slide).Replace("slide_", "")) - 1;
					presentation.Slides.RemoveAt(index);
					var inserted = presentation.Slides.InsertEmptySlide(index, presentation.LayoutSlides[0]);

					// temporary workaround - inserting SVG as an image, because import from SVG doesn't work now
					string svgContent = File.ReadAllText(slide);
					ISvgImage svgImage = new SvgImage(svgContent);
					inserted.Shapes.AddGroupShape(svgImage, 0, 0, presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
				}

				presentation.Save(outFile, Aspose.Slides.Export.SaveFormat.Pptx);
			}

			return Task.CompletedTask;
		}
	}
}
