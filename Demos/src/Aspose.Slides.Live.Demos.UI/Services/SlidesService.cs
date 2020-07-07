using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using Aspose.Slides.Live.Demos.UI.Models;
using System.Drawing.Imaging;

namespace Aspose.Slides.Live.Demos.UI.Services
{
	/// <summary>
	/// Class contains business logic for Slides.App.
	/// </summary>
	public partial class SlidesService : BaseService, IDisposable
	{
		/// <summary>
		/// Releases all resources used by this object.
		/// </summary>
		public void Dispose()
		{
			// intentionally blank
		}

		static SlidesService()
		{
			Models.License.SetAsposeSlidesLicense();
			Models.License.SetAsposeWordsLicense();
		}

		private Aspose.Slides.Export.SaveFormat GetFormatFromSource(string sourceFile)
		{
			switch (Path.GetExtension(sourceFile))
			{
				case ".ppt":
					return Aspose.Slides.Export.SaveFormat.Ppt;
				case ".odp":
					return Aspose.Slides.Export.SaveFormat.Odp;
				case ".pptx":
				default:
					return Aspose.Slides.Export.SaveFormat.Pptx;
			}
		}
		private ImageFormat GetImageFormat(SlidesConversionFormat format)
		{
			switch (format)
			{
				case SlidesConversionFormat.bmp:
					return ImageFormat.Bmp;
				case SlidesConversionFormat.jpeg:
					return ImageFormat.Jpeg;
				case SlidesConversionFormat.png:
					return ImageFormat.Png;
				case SlidesConversionFormat.emf:
					return ImageFormat.Wmf;
				case SlidesConversionFormat.wmf:
					return ImageFormat.Wmf;
				case SlidesConversionFormat.gif:
					return ImageFormat.Gif;
				case SlidesConversionFormat.exif:
					return ImageFormat.Emf;
				case SlidesConversionFormat.ico:
					return ImageFormat.Icon;
				default:
					throw new ArgumentException($"Unknown format {format}");
			}
		}
	}
}
