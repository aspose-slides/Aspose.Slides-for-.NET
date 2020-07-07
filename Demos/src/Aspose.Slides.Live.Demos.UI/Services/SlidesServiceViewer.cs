using System;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Live.Demos.UI.Models;
using Aspose.Slides;
using Aspose.Slides.Export;
using PresentationInfo = Aspose.Slides.Live.Demos.UI.Models.Slides.PresentationInfo;

namespace Aspose.Slides.Live.Demos.UI.Services
{
	public partial class SlidesService
	{
		private const int ThumbnailWidth = 150;
		private const int ThumbnailHeight = 150;

		public async Task PrepareViewerAsync(string source, string destination)
		{
			using (var sourceStream = File.Open(source, FileMode.Open))
			using (var destinationStream = File.Create(destination))
			{
				await sourceStream.CopyToAsync(destinationStream);
			}
		}

		public PresentationSlide[] GetThumbnails(string filePath)
		{
			using (var presentation = new Presentation(filePath))
			{
				return presentation.Slides.Select(slide =>
				{
					float scaleX = 1f / presentation.SlideSize.Size.Width * ThumbnailWidth;
					float scaleY = 1f / presentation.SlideSize.Size.Height * ThumbnailHeight;
					float scale = Math.Min(scaleX, scaleY);
					using (var stream = new MemoryStream())
					using (var bitmap = slide.GetThumbnail(scale, scale))
					{
						bitmap.Save(stream, ImageFormat.Png);
						return new PresentationSlide
						{
							Number = slide.SlideNumber,
							Height = bitmap.Height,
							Width = bitmap.Width,
							Angle = 0,
							Data = Convert.ToBase64String(stream.ToArray())
						};
					}
				}).ToArray();
			}
		}

		public PresentationSlide GetViewerSlide(string filePath, int slideNumber)
		{
			using (var presentation = new Presentation(filePath))
			using (var stream = new MemoryStream())
			{
				presentation.Save(stream, new[] { slideNumber }, SaveFormat.Html, new HtmlOptions
				{
					SvgResponsiveLayout = true,
					HtmlFormatter = HtmlFormatter.CreateCustomFormatter(new EmptyHtmlFormatter())
				});

				return new PresentationSlide
				{
					Number = slideNumber,
					Height = presentation.SlideSize.Size.Height,
					Width = presentation.SlideSize.Size.Width,
					Angle = 0,
					Data = new UTF8Encoding(false).GetString(stream.ToArray())
				};
			}
		}

		public Aspose.Slides.Live.Demos.UI.Models.Slides.PresentationInfo GetViewerInfo(string filePath)
		{
			using (var presentation = new Presentation(filePath))
			{
				return new Aspose.Slides.Live.Demos.UI.Models.Slides.PresentationInfo
				{
					Slides = presentation.Slides.Select(s => new PresentationSlide
					{
						Number = s.SlideNumber,
						Angle = 0,
						Height = presentation.SlideSize.Size.Height,
						Width = presentation.SlideSize.Size.Width
					}).ToList(),
					Guid = Path.GetFileName(filePath),
					NavigationItems = presentation.Slides.Select(s => new PresentationNavigationItem
					{
						Name = s.SlideNumber.ToString(),
						Number = s.SlideNumber,
						Style = 1 // Heading1 style
					}).ToList()
				};
			}
		}
	}
}
