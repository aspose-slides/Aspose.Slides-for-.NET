using Aspose.Slides.Export;
using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Web.Core.Enums;
using Aspose.Slides.Web.Core.Infrastructure;
using Aspose.Slides.Web.Interfaces.Services;
using Microsoft.Extensions.Logging;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading;

namespace Aspose.Slides.Web.Core.Services
{
	/// <summary>
	/// Implementation of signature service logic.
	/// </summary>
	internal sealed class SignatureService : SlidesServiceBase, ISignatureService
	{
		private readonly IConversionService _conversionService;

		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="conversionService"></param>
		/// <param name="licenseProvider"></param>
		public SignatureService(ILogger<SignatureService> logger, IConversionService conversionService, ILicenseProvider licenseProvider) : base(logger)
		{
			_conversionService = conversionService;
			licenseProvider.SetAsposeLicense(AsposeProducts.Slides);
		}

		/// <summary>
		///     Adds signature to the each presentation slide.
		/// </summary>
		/// <param name="inputFile">The presentation file.</param>
		/// <param name="destinationFolder">The output folder.</param>
		/// <param name="format">The output format.</param>
		/// <param name="image">The signature image stream. Should be null when the text signature is added.</param>
		/// <param name="text">The signature text.</param>
		/// <param name="color">The color of the signature text (ignored when image signature is added).</param>
		/// <param name="cancellationToken">Cancellation token.</param>
		/// <returns></returns>
		public IEnumerable<string> Sign(string inputFile, string destinationFolder, SlidesConversionFormats format, Stream image,
			string text, Color color, CancellationToken cancellationToken = default)
		{
			var resultFile = Path.Combine(destinationFolder, Path.GetFileName(inputFile));
			using (var p = new Presentation(inputFile))
			{
				PPImage ppImage = null;
				if (image != null)
					using (var bitmap = new Bitmap(image))
					{
						ppImage = (PPImage) p.Images.AddImage(bitmap);
					}

				const int margin = 10;
				const int textSignatureHeight = 50;
				const int textSignatureWidth = 150;
				foreach (var slide in p.Slides)
				{
					if (ppImage != null)
					{
						slide.Shapes.AddPictureFrame(
							ShapeType.Rectangle,
							p.SlideSize.Size.Width - ppImage.Width - margin,
							p.SlideSize.Size.Height - ppImage.Height - margin,
							ppImage.Width,
							ppImage.Height,
							ppImage
						);
					}
					else
					{
						var textShape = slide.Shapes.AddAutoShape(
							ShapeType.Rectangle,
							p.SlideSize.Size.Width - textSignatureWidth - margin,
							p.SlideSize.Size.Height - textSignatureHeight - margin,
							textSignatureWidth, textSignatureHeight
						);
						textShape.FillFormat.FillType = FillType.NoFill;
						textShape.LineFormat.FillFormat.FillType = FillType.NoFill;
						var portion = textShape.AddTextFrame(text).Paragraphs[0].Portions[0];
						portion.PortionFormat.FontItalic = NullableBool.True;
						portion.PortionFormat.FillFormat.FillType = FillType.Solid;
						portion.PortionFormat.FillFormat.SolidFillColor.Color = color;
					}

					cancellationToken.ThrowIfCancellationRequested();
				}

				p.Save(resultFile, SaveFormat.Pptx);
			}

			if (format != SlidesConversionFormats.pptx)
				return _conversionService.Conversion(
					new string[] { resultFile },
					destinationFolder,
					format,
					cancellationToken
				);

			return new List<string>() { resultFile };
		}
	}
}
