using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Export;
using System;
using System.IO;
using System.Drawing.Imaging;

namespace Aspose.Slides.Web.Core.Helpers
{
	/// <summary>
	/// The methods-extensions for convert of file formats
	/// </summary>
	internal static class FileFormatsConverterExtensions
	{
		/// <summary>
		/// Converts PresentationFormats to Aspose.Slides.Export.SaveFormat
		/// </summary>
		/// <param name="presentationFormat"></param>
		/// <returns></returns>
		public static SaveFormat ToSaveFormat(this PresentationFormats presentationFormat)
		{
			return (SaveFormat)Enum.Parse(typeof(SaveFormat), presentationFormat.ToString());
		}

		/// <summary>
		/// Converts the SlidesConversionFormats to the ImageFormat
		/// </summary>
		/// <param name="format">The SlidesConversionFormats</param>
		/// <returns>The ImageFormat</returns>
		public static ImageFormat GetImageFormat(this SlidesConversionFormats format)
		{
			switch (format)
			{
				case SlidesConversionFormats.bmp:
					return ImageFormat.Bmp;
				case SlidesConversionFormats.jpeg:
					return ImageFormat.Jpeg;
				case SlidesConversionFormats.png:
					return ImageFormat.Png;
				case SlidesConversionFormats.emf:
					return ImageFormat.Emf;
				case SlidesConversionFormats.wmf:
					return ImageFormat.Wmf;
				case SlidesConversionFormats.gif:
					return ImageFormat.Gif;
				case SlidesConversionFormats.exif:
					return ImageFormat.Exif;
				case SlidesConversionFormats.ico:
					return ImageFormat.Icon;
				default:
					throw new ArgumentException($"Unknown format {format}");
			}
		}

		/// <summary>
		/// Gets SaveFormat by presentation source file
		/// </summary>
		/// <param name="sourceFile">source file of presentation</param>
		/// <returns>Slides.Export.SaveFormat</returns>
		public static SaveFormat GetSlidesExportSaveFormatBySourceFile(this string sourceFile)
		{
			switch (Path.GetExtension(sourceFile).TrimStart('.').ToLowerInvariant())
			{				
				case "odp":
					return SaveFormat.Odp;
				case "otp":
					return SaveFormat.Otp;
				case "ppt":
					return SaveFormat.Ppt;
				case "potx":
					return SaveFormat.Potx;
				case "potm":
					return SaveFormat.Potm;
				case "pot":
					return SaveFormat.Pot;
				case "pptm":
					return SaveFormat.Pptm;
				case "ppsm":
					return SaveFormat.Ppsm;
				case "pps":
					return SaveFormat.Pps;
				case "pptx":
				default:
					return SaveFormat.Pptx;
			}
		}
	}
}
