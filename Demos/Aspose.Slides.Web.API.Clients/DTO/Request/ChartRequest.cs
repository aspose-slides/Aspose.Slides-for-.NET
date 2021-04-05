using Aspose.Slides.Web.API.Clients.Enums;

namespace Aspose.Slides.Web.API.Clients.DTO.Request
{
	/// <summary>
	/// The request model for the charts
	/// </summary>
	public sealed class ChartRequest : BaseRequest
	{
		/// <summary>
		/// The type of a chart
		/// </summary>
		public ChartTypes ChartType { get; set; }

		/// <summary>
		/// The file format for saving
		/// </summary>
		public SlidesConversionFormats SaveFormat { get; set; }

		/// <summary>
		/// Is input data in the request 
		/// </summary>
		public bool IsExternalData { get; set; } = false;

		/// <summary>
		/// The external data field
		/// </summary>
		public string JsonData { get; set; }

		/// <summary>
		/// Attribute is full size or preview chart
		/// </summary>
		public bool IsPreview { get; set; } = false;
	}
}
