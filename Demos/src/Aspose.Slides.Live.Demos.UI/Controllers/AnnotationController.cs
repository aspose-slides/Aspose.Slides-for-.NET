using Aspose.Slides.Live.Demos.UI.Models.Common;
using Aspose.Slides.Live.Demos.UI.Models;
using Aspose.Slides.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;

namespace Aspose.Slides.Live.Demos.UI.Controllers
{
	public class AnnotationController : BaseController  
	{
		public override string Product => (string)RouteData.Values["product"];


		[HttpPost]
		public Response Remove()
		{
			Response response = null;
			if (Request.Files.Count > 0)
			{
				string _sourceFolder = Guid.NewGuid().ToString();
				var docs = UploadFiles(Request, _sourceFolder);

				if (docs.Count > 0)
				{
					

					AsposeSlides asposeSlides = new AsposeSlides();
					FileSafeResult FileSafeResult = asposeSlides.RemoveAnnotations(docs[0].FolderName, docs[0].FileName);

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

		

		public ActionResult Annotation()
		{
			var model = new ViewModel(this, "Annotation")
			{
				
				MaximumUploadFiles = 1,
				
				DropOrUploadFileLabel = Resources["DropOrUploadFile"]
			};
			return View(model);
		}
		

	}
}
