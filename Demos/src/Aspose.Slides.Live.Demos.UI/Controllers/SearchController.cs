using Aspose.Slides.Live.Demos.UI.Models.Common;
using Aspose.Slides.Live.Demos.UI.Models;
using Aspose.Slides.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;

namespace Aspose.Slides.Live.Demos.UI.Controllers
{
	public class SearchController : BaseController
	{
		public override string Product => (string)RouteData.Values["product"];


		[HttpPost]
		public Response Search( string query)
		{
			Response response = null;
			if (Request.Files.Count > 0)
			{
				string _sourceFolder = Guid.NewGuid().ToString();
				var docs = UploadFiles(Request, _sourceFolder);

				if (docs.Count > 0)
				{
					SearchRequestModel searchRequestModel = new SearchRequestModel();
					searchRequestModel.id = docs[0].FolderName;
					searchRequestModel.FileName = docs[0].FileName;
					searchRequestModel.Query = query;

					AsposeSlides asposeSlides = new AsposeSlides();
					FileSafeResult FileSafeResult = asposeSlides.Search(searchRequestModel);

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
							Text = "No Search Result Found",
							FileProcessingErrorCode = FileProcessingErrorCode.NoSearchResults
						};
					}
				}

			}

			return response;
		}
		public ActionResult Search()
		{
			var model = new ViewModel(this, "Search")
			{
				ControlsView = "SearchControls",
				
				MaximumUploadFiles = 1,
				DropOrUploadFileLabel = Resources["DropOrUploadFile"]
			};
			if (model.RedirectToMainApp)
				return Redirect("/slides/" + model.AppName.ToLower());
			return View(model);
		}

	}
}
