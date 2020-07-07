using Aspose.Slides.Live.Demos.UI.Models;
using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Web;

using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Windows.Forms;


namespace Aspose.Slides.Live.Demos.UI.Services
{
	public partial class SlidesService
	{
		/// <summary>
		/// Adds text watermark into source file, saves resulted file to out file.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFile">Output slides file.</param>
		/// <param name="options">Watermark options.</param>
		public void AddTextWatermark(
			string sourceFile,
			string outFile,
			TextWatermarkOptionsModel options
		)
		{
			using (var presentation = new Presentation(sourceFile))
			{
				var size = presentation.SlideSize.Size;
				var height = size.Height;
				var width = size.Width;

				var centerW = (size.Width - width) / 2;
				var centerH = (size.Height - height) / 2;

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
				}

				presentation.Save(outFile, GetFormatFromSource(sourceFile));
			}
		}
		/// <summary>
		/// ResizeImage
		/// </summary>
		private  Bitmap ResizeImage(Image image, double zoom)
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
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFile">Output slides file.</param>
		/// <param name="options">Watermark options.</param>		
		public void AddImageWatermark(
			string sourceFile,
			string outFile,
			ImageWatermarkOptionsModel options
		)
		{
			using (var presentation = new Presentation(sourceFile))
			{
				var size = presentation.SlideSize.Size;
				var img = ResizeImage(
					new Bitmap(options.ImageFile),
					options.ZoomPercent / 100.0
				) as Image;

				if (options.IsGrayScaled)
					img = ToolStripRenderer.CreateDisabledImage(img);

				var imgx = presentation.Images.AddImage(img);

				var height = imgx.Height;
				var width = imgx.Width;

				var centerW = (size.Width - width) / 2;
				var centerH = (size.Height - height) / 2;

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
				}

				presentation.Save(outFile, GetFormatFromSource(sourceFile));
			}
		}

		/// <summary>
		/// Removes watermark from source file, saves resulted file to out file.		
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFile">Output slides file.</param>		
		public void RemoveWatermark(
			string sourceFile,
			string outFile
		)
		{
			using (var presentation = new Presentation(sourceFile))
			{
				foreach (var slide in presentation.Slides)
				{
					var watermarks = slide.Shapes.Where(s =>
						s.Name.Contains("WaterMark")
						|| (s as AutoShape)?.TextFrame?.Text?.Contains("WaterMark") == true
					).ToList();

					foreach (var shape in watermarks)
						slide.Shapes.Remove(shape);
				}

				presentation.Save(outFile, GetFormatFromSource(sourceFile));
			}
		}

		

		

		
	}
}
