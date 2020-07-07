using Aspose.Slides.Live.Demos.UI.Models.Common;
using Aspose.Slides.Live.Demos.UI.Models;
using Aspose.Slides.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;

namespace Aspose.Slides.Live.Demos.UI.Controllers
{
	public class RedactionController : BaseController  
	{
		public override string Product => (string)RouteData.Values["product"];


		[HttpPost]
		public Response Redaction(string outputType, string searchQuery, string replaceText, bool caseSensitive, bool text, bool comments, bool metadata)
		{
			Response response = null;
			if (Request.Files.Count > 0)
			{
				string _sourceFolder = Guid.NewGuid().ToString();
				var docs = UploadFiles(Request, _sourceFolder);

				if (docs.Count > 0)
				{
					
					AsposeSlides asposeSlides = new AsposeSlides();

					RedactionOptionsModel redactionOptionsModel = new RedactionOptionsModel();
					redactionOptionsModel.id = docs[0].FolderName;
					redactionOptionsModel.FileName = docs[0].FileName;
					redactionOptionsModel.ReplaceText = replaceText;
					redactionOptionsModel.SearchQuery = searchQuery;
					redactionOptionsModel.IsCaseSensitiveSearch = caseSensitive;
					redactionOptionsModel.MustReplaceText = text;
					redactionOptionsModel.MustReplaceComments = comments;
					redactionOptionsModel.MustReplaceMetadata = metadata;

					FileSafeResult FileSafeResult = asposeSlides.Redaction(redactionOptionsModel);

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
		public ActionResult Redaction()
		{
			var model = new ViewModel(this, "Redaction")
			{
				ControlsView = "RedactionControls",
				
				MaximumUploadFiles = 1,
				DropOrUploadFileLabel = Resources["DropOrUploadFile"]
			};
			if (model.RedirectToMainApp)
				return Redirect("/slides/" + model.AppName.ToLower());
			return View(model);			
		}	

	}
}
