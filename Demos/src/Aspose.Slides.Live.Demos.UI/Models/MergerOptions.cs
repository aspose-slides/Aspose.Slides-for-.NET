using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Aspose.Slides.Live.Demos.UI.Models
{
	/// <summary>
	/// Merger options.
	/// </summary>
	public class MergerOptions
	{
		/// <summary>
		/// Upload id.
		/// </summary>
		public string idMain { get; set; }

		public string[] MainFiles { get; set; }

		/// <summary>
		/// Upload id for StyleMaster.
		/// </summary>
		public string idStyleMaster { get; set; }

		/// <summary>
		/// File name for StyleMaster.
		/// </summary>
		public string FileNameStyleMaster { get; set; }

		/// <summary>
		/// Format to convert into.
		/// When not specified, used pptx.
		/// </summary>
		public string Format { get; set; }
	}
}
