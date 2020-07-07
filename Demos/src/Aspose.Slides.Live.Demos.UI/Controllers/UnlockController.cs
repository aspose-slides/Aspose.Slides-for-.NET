using Aspose.Slides.Live.Demos.UI.Models.Common;
using Aspose.Slides.Live.Demos.UI.Models;
using Aspose.Slides.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;

namespace Aspose.Slides.Live.Demos.UI.Controllers
{
	public class UnlockController : BaseController  
	{
		public override string Product => (string)RouteData.Values["product"];


		[HttpPost]
		public Response Unlock( string passw)
		{
			Response response = null;
			if (Request.Files.Count > 0)
			{
				string _sourceFolder = Guid.NewGuid().ToString();
				var docs = UploadFiles(Request, _sourceFolder);

				if (docs.Count > 0)
				{
					UnProtectOptions unProtectOptions = new UnProtectOptions();
					unProtectOptions.id = docs[0].FolderName;
					unProtectOptions.FileName = docs[0].FileName;
					unProtectOptions.Password =passw;

					AsposeSlides asposeSlides = new AsposeSlides();
					FileSafeResult FileSafeResult = asposeSlides.Unlock(unProtectOptions);

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
		public ActionResult Unlock()
		{
			var model = new ViewModel(this, "Unlock")
			{
				ControlsView = "UnlockControls",
				SaveAsComponent = true,
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
