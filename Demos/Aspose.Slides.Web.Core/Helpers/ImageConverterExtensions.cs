using System;
using System.Drawing;
using System.Drawing.Imaging;

namespace Aspose.Slides.Web.Core.Helpers
{
	internal static class ImageConverterExtensions
	{
		// See https://stackoverflow.com/questions/2265910/convert-an-image-to-grayscale
		internal static Bitmap ConvertToGrayscale(this Bitmap original)
		{
			//create a blank bitmap the same size as original
			Bitmap newBitmap = new Bitmap(original.Width, original.Height);

			//get a graphics object from the new image
			using Graphics g = Graphics.FromImage(newBitmap);

			//create the grayscale ColorMatrix
			ColorMatrix colorMatrix = new ColorMatrix(
				new float[][]
				{
					new float[] {.3f, .3f, .3f, 0, 0},
					new float[] {.59f, .59f, .59f, 0, 0},
					new float[] {.11f, .11f, .11f, 0, 0},
					new float[] {0, 0, 0, 1, 0},
					new float[] {0, 0, 0, 0, 1}
				});

			//create some image attributes
			using ImageAttributes attributes = new ImageAttributes();
			
			//set the color matrix attribute
			attributes.SetColorMatrix(colorMatrix);

			//draw the original image on the new image
			//using the grayscale color matrix
			g.DrawImage(original, new Rectangle(0, 0, original.Width, original.Height),
						0, 0, original.Width, original.Height, GraphicsUnit.Pixel, attributes);
			
			return newBitmap;
		}

		// https://stackoverflow.com/questions/1940581/c-sharp-image-resizing-to-different-size-while-preserving-aspect-ratio
		internal static SizeF ResizeKeepAspect(this SizeF src, float maxWidth, float maxHeight, bool enlarge = false)
		{
			maxWidth = enlarge ? maxWidth : Math.Min(maxWidth, src.Width);
			maxHeight = enlarge ? maxHeight : Math.Min(maxHeight, src.Height);

			var scale = Math.Min(maxWidth / src.Width, maxHeight / src.Height);
			return new SizeF((float)Math.Round(src.Width * scale), (float)Math.Round(src.Height * scale));
		}
	}
}
