using Aspose.Slides.Web.Core.Enums;
using Aspose.Slides.Web.Core.Helpers;
using Aspose.Slides.Web.Core.Infrastructure;
using Aspose.Slides.Web.Interfaces.Models.Watermark;
using Aspose.Slides.Web.Interfaces.Services;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Core.Services
{
	/// <summary>
	/// Implementation of slides watermark logic.
	/// </summary>
	internal sealed class WatermarkService : SlidesServiceBase, IWatermarkService
	{
		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="licenseProvider"></param>
		public WatermarkService(ILogger<WatermarkService> logger, ILicenseProvider licenseProvider) : base(logger)
		{
			licenseProvider.SetAsposeLicense(AsposeProducts.Slides);
		}

		/// <summary>
		/// Adds text watermark into source files, saves resulted files to out files.
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFiles">Output slides files.</param>
		/// <param name="options">Watermark options.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		public void AddTextWatermark(
			IList<string> sourceFiles,
			IList<string> outFiles,
			TextWatermarkOptions options,
			CancellationToken cancellationToken = default
		)
		{
			void AddTextWatermarkByIndex(int index)
			{
				var sourceFile = sourceFiles[index];
				using var presentation = new Presentation(sourceFile);
				
				cancellationToken.ThrowIfCancellationRequested();

				var size = presentation.SlideSize.Size;
				var height = size.Height;
				var width = size.Width;
				var centerW = (size.Width - width) / 2;
				var centerH = (size.Height - height) / 2;

				cancellationToken.ThrowIfCancellationRequested();

				foreach (var slide in presentation.Slides)
				{
					var shape = slide.Shapes.AddAutoShape(
						ShapeType.Rectangle,
						centerW, centerH,
						width, height
					);
					shape.Name = "WaterMark";

					shape.FillFormat.FillType = FillType.NoFill;
					shape.LineFormat.FillFormat.FillType = FillType.NoFill;

					var textFrame = shape.AddTextFrame(" ");
					textFrame.TextFrameFormat.AnchoringType = TextAnchorType.Center;
					textFrame.TextFrameFormat.CenterText = NullableBool.True;

					var paragraph = textFrame.Paragraphs[0];
					paragraph.ParagraphFormat.Alignment = TextAlignment.Center;

					var portion = paragraph.Portions[0];
					var format = portion.PortionFormat;

					format.FillFormat.FillType = FillType.Solid;

					portion.Text = options.Text;
					format.FillFormat.SolidFillColor.Color = options.ColorValue;
					format.LatinFont = new FontData(options.FontName);
					format.FontHeight = options.FontSize;
					shape.Rotation = options.RotationAngleDegrees;

					cancellationToken.ThrowIfCancellationRequested();
				}

				cancellationToken.ThrowIfCancellationRequested();
				
				var outFile = outFiles[index];

				presentation.Save(outFile, sourceFile.GetSlidesExportSaveFormatBySourceFile());
			}

			try
			{
				Parallel.For(0, sourceFiles.Count, AddTextWatermarkByIndex);
			}
			catch (AggregateException ae)
			{
				foreach (var e in ae.InnerExceptions)
				{
					throw e;
				}
			}
		}

		/// <summary>
		/// ResizeImage
		/// </summary>
		Bitmap ResizeImage(Image image, double zoom)
		{
			var width = (int)(image.Width * zoom);
			var height = (int)(image.Height * zoom);

			var destRect = new Rectangle(0, 0, width, height);
			var destImage = new Bitmap(width, height);

			destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

			using (var graphics = Graphics.FromImage(destImage))
			{
				graphics.CompositingMode = CompositingMode.SourceCopy;
				graphics.CompositingQuality = CompositingQuality.HighQuality;
				graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
				graphics.SmoothingMode = SmoothingMode.HighQuality;
				graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

				using (var wrapMode = new ImageAttributes())
				{
					wrapMode.SetWrapMode(WrapMode.TileFlipXY);
					graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
				}
			}

			return destImage;
		}

		/// <summary>
		/// Adds image watermark into source file, saves resulted file to out file.
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFiles">Output slides files.</param>
		/// <param name="options">Watermark options.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		public void AddImageWatermark(
			IList<string> sourceFiles,
			IList<string> outFiles,
			ImageWatermarkOptions options,
			CancellationToken cancellationToken = default
		)
		{
			void AddImageWatermarkByIndex(int index)
			{
				var sourceFile = sourceFiles[index];
				using var presentation = new Presentation(sourceFile);

				cancellationToken.ThrowIfCancellationRequested();

				var size = presentation.SlideSize.Size;
				using var bitmap = new Bitmap(options.ImageFile);
				var img = ResizeImage(
					bitmap,
					options.ZoomPercent / 100.0
				);

				if (options.IsGrayScaled)
				{
					img = img.ConvertToGrayscale();
				}

				var imgx = presentation.Images.AddImage(img);

				var height = imgx.Height;
				var width = imgx.Width;

				var centerW = (size.Width - width) / 2;
				var centerH = (size.Height - height) / 2;

				cancellationToken.ThrowIfCancellationRequested();
				foreach (var slide in presentation.Slides)
				{
					IPictureFrame pf = slide.Shapes.AddPictureFrame(
						ShapeType.Rectangle,
						centerW, centerH,
						imgx.Width, imgx.Height,
						imgx
					);

					pf.Name = "WaterMark";

					pf.LineFormat.FillFormat.FillType = FillType.NoFill;

					pf.Rotation = options.RotationAngleDegrees;

					cancellationToken.ThrowIfCancellationRequested();
				}

				cancellationToken.ThrowIfCancellationRequested();

				var outFile = outFiles[index];
				presentation.Save(outFile, sourceFile.GetSlidesExportSaveFormatBySourceFile());
			}

			try
			{
				Parallel.For(0, sourceFiles.Count, AddImageWatermarkByIndex);
			}
			catch (AggregateException ae)
			{
				foreach (var e in ae.InnerExceptions)
				{
					throw e;
				}
			}
		}

		/// <summary>
		/// Removes watermark from source file, saves resulted file to out file.
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFiles">Output slides files.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		public void RemoveWatermark(
			IList<string> sourceFiles,
			IList<string> outFiles,
			CancellationToken cancellationToken = default
		)
		{
			void RemoveWatermarkByIndex(int index)
			{
				var sourceFile = sourceFiles[index];
				using var presentation = new Presentation(sourceFile);

				cancellationToken.ThrowIfCancellationRequested();

				foreach (var slide in presentation.Slides)
				{
					var watermarks = slide.Shapes.Where(s =>
						s.Name.Contains("WaterMark")
						|| (s as AutoShape)?.TextFrame?.Text?.Contains("WaterMark") == true
					).ToList();

					cancellationToken.ThrowIfCancellationRequested();
					foreach (var shape in watermarks)
					{
						slide.Shapes.Remove(shape);
						cancellationToken.ThrowIfCancellationRequested();
					}

					cancellationToken.ThrowIfCancellationRequested();
				}

				cancellationToken.ThrowIfCancellationRequested();

				var outFile = outFiles[index];
				presentation.Save(outFile, sourceFile.GetSlidesExportSaveFormatBySourceFile());
			}

			try
			{
				Parallel.For(0, sourceFiles.Count, RemoveWatermarkByIndex);
			}
			catch (AggregateException ae)
			{
				foreach (var e in ae.InnerExceptions)
				{
					throw e;
				}
			}
		}

		/// <summary>
		/// Asynchronously adds text watermark into source files, saves resulted files to out files.
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFiles">Output slides files.</param>
		/// <param name="options">Watermark options.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		public async Task AddTextWatermarkAsync(
			IList<string> sourceFiles,
			IList<string> outFiles,
			TextWatermarkOptions options,
			CancellationToken cancellationToken = default
		) => await Task.Run(() => AddTextWatermark(sourceFiles, outFiles, options, cancellationToken));

		/// <summary>
		/// Asynchronously adds image watermark into source file, saves resulted file to out file.
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFiles">Output slides files.</param>
		/// <param name="options">Watermark options.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		public async Task AddImageWatermarkAsync(
			IList<string> sourceFiles,
			IList<string> outFiles,
			ImageWatermarkOptions options,
			CancellationToken cancellationToken = default
		) => await Task.Run(() => AddImageWatermark(sourceFiles, outFiles, options, cancellationToken));

		/// <summary>
		/// Asynchronously removes watermark from source file, saves resulted file to out file.		
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFiles">Output slides files.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		public async Task RemoveWatermarkAsync(
			IList<string> sourceFiles,
			IList<string> outFiles,
			CancellationToken cancellationToken = default
		) => await Task.Run(() => RemoveWatermark(sourceFiles, outFiles, cancellationToken));
	}
}
