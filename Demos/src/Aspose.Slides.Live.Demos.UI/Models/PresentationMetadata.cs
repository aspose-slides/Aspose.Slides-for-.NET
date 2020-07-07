using System;
using System.Collections.Generic;

namespace Aspose.Slides.Live.Demos.UI.Models
{
	/// <summary>
	/// Represents presentation document metadata.
	/// </summary>
	public class PresentationMetadata
	{
		/// <summary>
		/// Returns the app version.
		/// Read-only <see cref="T:System.String" />.
		/// </summary>
		public string AppVersion { get; set; }

		/// <summary>
		/// Returns or sets the name of the application.
		/// Read/write <see cref="T:System.String" />.
		/// </summary>
		public string NameOfApplication { get; set; }

		/// <summary>
		/// Returns or sets the company property.
		/// Read/write <see cref="T:System.String" />.
		/// </summary>
		public string Company { get; set; }

		/// <summary>
		/// Returns or sets the manager property.
		/// Read/write <see cref="T:System.String" />.
		/// </summary>
		public string Manager { get; set; }

		/// <summary>
		/// Returns or sets the intended format of a presentation.
		/// Read/write <see cref="T:System.String" />.
		/// </summary>
		public string PresentationFormat { get; set; }

		/// <summary>
		/// Determines whether the presentation is shared between multiple people.
		/// Read/write <see cref="T:System.Boolean" />.
		/// </summary>
		public bool SharedDoc { get; set; }

		/// <summary>
		/// Returns or sets the template of a application.
		/// Read/write <see cref="T:System.String" />.
		/// </summary>
		public string ApplicationTemplate { get; set; }

		/// <summary>
		/// Total editing time of a presentation.
		/// Read/write <see cref="T:System.TimeSpan" />.
		/// </summary>
		public TimeSpan TotalEditingTime { get; set; }

		/// <summary>
		/// Returns or sets the title of a presentation.
		/// Read/write <see cref="T:System.String" />.
		/// </summary>
		public string Title { get; set; }

		/// <summary>
		/// Returns or sets the subject of a presentation.
		/// Read/write <see cref="T:System.String" />.
		/// </summary>
		public string Subject { get; set; }

		/// <summary>
		/// Returns or sets the author of a presentation.
		/// Read/write <see cref="T:System.String" />.
		/// </summary>
		public string Author { get; set; }

		/// <summary>
		/// Returns or sets the keywords of a presentation.
		/// Read/write <see cref="T:System.String" />.
		/// </summary>
		public string Keywords { get; set; }

		/// <summary>
		/// Returns or sets the comments of a presentation.
		/// Read/write <see cref="T:System.String" />.
		/// </summary>
		public string Comments { get; set; }

		/// <summary>
		/// Returns or sets the category of a presentation.
		/// Read/write <see cref="T:System.String" />.
		/// </summary>
		public string Category { get; set; }

		/// <summary>
		/// Returns the date when a presentation was created.
		/// Read/write <see cref="T:System.DateTime" />.
		/// </summary>
		public DateTime CreatedTime { get; set; }

		/// <summary>
		/// Returns the date when a presentation was modified last time.
		/// Read-only in case of Presentation.DocumentProperties (because it will be updated internally while IPresentation object saving process).
		/// Can be changed via DocumentProperties instance returning by method <see cref="M:Aspose.Slides.IPresentationInfo.ReadDocumentProperties" />
		/// Please see the example in <see cref="M:Aspose.Slides.IPresentationInfo.UpdateDocumentProperties(Aspose.Slides.IDocumentProperties)" /> method summary.
		/// </summary>
		public DateTime LastSavedTime { get; set; }

		/// <summary>
		/// Returns the date when a presentation was printed last time.
		/// Read/write <see cref="T:System.DateTime" />.
		/// </summary>
		public DateTime LastPrinted { get; set; }

		/// <summary>
		/// Returns or sets the name of a last person who modified a presentation.
		/// Read/write <see cref="T:System.String" />.
		/// </summary>
		public string LastSavedBy { get; set; }

		/// <summary>
		/// Returns or sets the presentation revision number.
		/// Read/write <see cref="T:System.Int32" />.
		/// </summary>
		public int RevisionNumber { get; set; }

		/// <summary>
		/// Returns or sets the content status of a presentation.
		/// Read/write <see cref="T:System.String" />.
		/// </summary>
		public string ContentStatus { get; set; }

		/// <summary>
		/// Returns or sets the content type of a presentation.
		/// Read/write <see cref="T:System.String" />.
		/// </summary>
		public string ContentType { get; set; }

		/// <summary>
		/// Returns or sets the HyperlinkBase document property.
		/// Read/write <see cref="T:System.String" />.
		/// </summary>
		public string HyperlinkBase { get; set; }

		/// <summary>
		/// Returns or sets the custom properties of the document.
		/// </summary>
		public Dictionary<string, object> CustomProperties { get; set; } = new Dictionary<string, object>();
		
	}
	
}
