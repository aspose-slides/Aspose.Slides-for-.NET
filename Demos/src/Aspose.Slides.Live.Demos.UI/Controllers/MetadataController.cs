using Aspose.Slides.Live.Demos.UI.Models.Common;
using Aspose.Slides.Live.Demos.UI.Models;
using Aspose.Slides.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;

namespace Aspose.Slides.Live.Demos.UI.Controllers
{
	public class MetadataController : BaseController
	{
		public override string Product => (string)RouteData.Values["product"];

		public ActionResult Metadata()
		{
			var model = new ViewModel(this, "Metadata")
			{
				ControlsView = "MetadataControls",
				UploadAndRedirect = true,
				MaximumUploadFiles = 1,
				DropOrUploadFileLabel = Resources["DropOrUploadFile"]
			};
			if (model.RedirectToMainApp)
				return Redirect("/slides/" + model.AppName.ToLower());
			return View(model);
		}

	}
}
