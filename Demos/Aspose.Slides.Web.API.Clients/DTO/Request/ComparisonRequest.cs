using Aspose.Slides.Web.API.Clients.Enums;

namespace Aspose.Slides.Web.API.Clients.DTO.Request
{
	/// <summary>
	/// The request model for the comparison
	/// </summary>
	public sealed class ComparisonRequest : BaseRequest
	{
		/// <summary>
		/// The comparison methods
		/// </summary>
		public ComparisonMethods ComparisonMethod { get; set; }

		/// <summary>
		/// The file format for saving of diff file
		/// </summary>
		public ComparisonDiffFileSaveFormats SaveFormat { get; set; }

		/// <summary>
		/// Folder name of the first file for compare.
		/// </summary>
		public string FirstFolderId { get; set; }

		/// <summary>
		/// File name of the first file for compare.
		/// </summary>
		public string FirstFileName { get; set; }

		/// <summary>
		/// Folder name of the second file for compare.
		/// </summary>
		public string SecondFolderId { get; set; }

		/// <summary>
		/// File name of the second file for compare.
		/// </summary>
		public string SecondFileName { get; set; }
	}
}
