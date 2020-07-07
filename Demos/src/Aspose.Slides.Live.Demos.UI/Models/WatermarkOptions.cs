using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;

namespace Aspose.Slides.Live.Demos.UI.Models
{
	/// <summary>
	/// Text watermark options.
	/// </summary>
	public class TextWatermarkOptionsModel : BaseRequestModel
	{
		/// <summary>
		/// Text.
		/// </summary>
		public string Text { get; set; }
		/// <summary>
		/// Color.
		/// </summary>
		public string Color { get; set; }
		/// <summary>
		/// Font name.
		/// </summary>
		public string FontName { get; set; }
		/// <summary>
		/// Font size.
		/// </summary>
		public int FontSize { get; set; }
		/// <summary>
		/// Rotation angle in degrees.
		/// </summary>
		public int RotationAngleDegrees { get; set; }

		internal double RotationAngleRadians => (Math.PI / 180) * RotationAngleDegrees;

		internal Color ColorValue
		{
			get
			{
				var colorString = Color;
				if (string.IsNullOrEmpty(colorString))
					colorString = "#FF808080"; // Gray

				return ColorTranslator.FromHtml(
					colorString.StartsWith("#") ? colorString : "#" + colorString
				);
			}
		}
	}

	/// <summary>
	/// Image watermark options.
	/// </summary>
	public class ImageWatermarkOptionsModel : BaseRequestModel
	{
		/// <summary>
		/// Upload id.
		/// </summary>
		public string idMain { get; set; }
		/// <summary>
		/// File name.
		/// </summary>
		public string MainFileName { get; set; }
		/// <summary>
		/// Is watermark must be gray scaled.
		/// </summary>
		public bool IsGrayScaled { get; set; }
		/// <summary>
		/// Zoom in percents.
		/// </summary>
		public int ZoomPercent { get; set; }
		/// <summary>
		/// Rotation angle in degrees.
		/// </summary>
		public int RotationAngleDegrees { get; set; }

		internal double RotationAngleRadians => (Math.PI / 180) * RotationAngleDegrees;

		internal string ImageFile { get; set; }
		/// <summary>
		/// Sets image file path.
		/// For tests.
		/// </summary>
		/// <param name="value">Path.</param>
		public void SetImageFile(string value) => ImageFile = value;
	}
}
