using System;
using System.Drawing;

namespace Aspose.Slides.Web.API.Clients.DTO.Request
{
	/// <summary>
	/// Text watermark options.
	/// </summary>
	public class TextWatermarkOptionsRequest : BaseRequest
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

		public double RotationAngleRadians => (Math.PI / 180) * RotationAngleDegrees;

		public Color ColorValue
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
	
}
