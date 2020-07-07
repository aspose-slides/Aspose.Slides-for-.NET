using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Aspose.Slides.Live.Demos.UI.Models
{
	/// <summary>
	/// UnProtectOptions request model.
	/// </summary>
	public class UnProtectOptions : BaseRequestModel
	{
		/// <summary>
		/// Password.
		/// </summary>
		public string Password { get; set; }
	}

	/// <summary>
	/// ProtectModel request model.
	/// </summary>
	public class ProtectOptions : BaseRequestModel
	{
		/// <summary>
		/// Password for view.
		/// </summary>
		public string PasswordView { get; set; }

		/// <summary>
		/// Password for edit.
		/// </summary>
		public string PasswordEdit { get; set; }

		/// <summary>
		/// Mark file as read-only.
		/// </summary>
		public bool MarkAsReadonly { get; set; }

		/// <summary>
		/// Mark file as final.
		/// </summary>
		public bool MarkAsFinal { get; set; }
	}
}
