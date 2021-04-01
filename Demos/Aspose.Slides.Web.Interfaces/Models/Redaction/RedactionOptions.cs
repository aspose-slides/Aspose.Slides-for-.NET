namespace Aspose.Slides.Web.Interfaces.Models.Redaction
{
	/// <summary>
	/// Redaction options.
	/// </summary>
	public sealed class RedactionOptions
	{
		/// <summary>
		/// Search query string.
		/// </summary>
		public string SearchQuery { get; set; }
		/// <summary>
		/// Replace string.
		/// </summary>
		public string ReplaceText { get; set; }
		/// <summary>
		/// Search is case sensitive.
		/// </summary>
		public bool IsCaseSensitiveSearch { get; set; }
		/// <summary>
		/// Replace text.
		/// </summary>
		public bool MustReplaceText { get; set; }
		/// <summary>
		/// Replace commentaries.
		/// </summary>
		public bool MustReplaceComments { get; set; }
		/// <summary>
		/// Replace metadata.
		/// </summary>
		public bool MustReplaceMetadata { get; set; }
	}
}
