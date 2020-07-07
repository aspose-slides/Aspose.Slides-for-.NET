using Aspose.Slides.Live.Demos.UI.Models.Common;
using Aspose.Slides.Live.Demos.UI.Models;
using Aspose.Slides.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;

namespace Aspose.Slides.Live.Demos.UI.Controllers
{
	public class LockController : BaseController  
	{
		public override string Product => (string)RouteData.Values["product"];


		[HttpPost]
		public Response Lock(string passw, string viewpassw, string makefinal)
		{
			Response response = null;
			if (Request.Files.Count > 0)
			{
				string _sourceFolder = Guid.NewGuid().ToString();
				var docs = UploadFiles(Request, _sourceFolder);

				if (docs.Count > 0)
				{
					ProtectOptions protectOptions = new ProtectOptions();
					protectOptions.id = docs[0].FolderName;
					protectOptions.FileName = docs[0].FileName;
					protectOptions.MarkAsFinal = bool.Parse(makefinal);
					protectOptions.MarkAsReadonly = false;
					protectOptions.PasswordEdit = passw;
					protectOptions.PasswordView = viewpassw;

					AsposeSlides asposeSlides = new AsposeSlides();
					FileSafeResult FileSafeResult = asposeSlides.Lock(protectOptions);

					if (FileSafeResult.IsSuccess)
					{
						return new Response
						{
							FileName = FileSafeResult.FileName,
							FolderName = FileSafeResult.id,
							StatusCode = 200,
							Text = "OK",
							FileProcessingErrorCode = FileProcessingErrorCode.OK
						};

					}
					else
					{
						return new Response
						{

							StatusCode = 500,
							Text = "Failed",
							FileProcessingErrorCode = FileProcessingErrorCode.OK
						};
					}
				}

			}

			return response;
		}
		public ActionResult Lock()
		{
			var model = new ViewModel(this, "Lock")
			{
				ControlsView = "LockControls",				
				MaximumUploadFiles = 1,
				DropOrUploadFileLabel = Resources["DropOrUploadFile"],
				ShowViewerButton = false
			};
			if (model.RedirectToMainApp)
				return Redirect("/slides/" + model.AppName.ToLower());
			return View(model);			
		}	

	}
}
