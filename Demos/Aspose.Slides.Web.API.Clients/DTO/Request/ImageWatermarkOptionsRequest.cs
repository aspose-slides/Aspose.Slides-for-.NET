using System;
using System.Collections.Generic;

namespace Aspose.Slides.Web.API.Clients.DTO.Request
{
	/// <summary>
	/// Image watermark options.
	/// </summary>
	public sealed class ImageWatermarkOptionsRequest : BaseRequest
	{
		/// <summary>
		/// Upload id.
		/// </summary>
		public string idMain { get; set; }
		/// <summary>
		/// File name.
		/// </summary>
		public IList<string> MainFileNames { get; set; }
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

		public double RotationAngleRadians => (Math.PI / 180) * RotationAngleDegrees;

		public string ImageFile { get; set; }
		/// <summary>
		/// Sets image file path.
		/// For tests.
		/// </summary>
		/// <param name="value">Path.</param>
		public void SetImageFile(string value) => ImageFile = value;
	}
}
