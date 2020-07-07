using Aspose.Slides.Live.Demos.UI.Models.Common;
using Aspose.Slides.Live.Demos.UI.Models;
using Aspose.Slides.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;


namespace Aspose.Slides.Live.Demos.UI.Controllers
{
	public class SplitterController : BaseController  
	{
		public override string Product => (string)RouteData.Values["product"];


		[HttpPost]
		public Response Splitter(string outputType, string splitType, string pars)
		{
			
			Response response = null;
			if (Request.Files.Count > 0)
			{
				string _sourceFolder = Guid.NewGuid().ToString();
				var docs = UploadFiles(Request, _sourceFolder);
				SplitType _splitType;
				Enum.TryParse((int.Parse(splitType) - 1).ToString(), out _splitType);
				if (docs.Count > 0)
				{
					SplitterRequestModel splitterRequestModel = new SplitterRequestModel();
					splitterRequestModel.id = docs[0].FolderName;
					splitterRequestModel.FileName = docs[0].FileName;
					splitterRequestModel.Format = outputType.Trim();
					splitterRequestModel.SplitType = _splitType;

					if (_splitType == SplitType.Range)
					{
						splitterRequestModel.SplitRange = pars;
					}
					else if (_splitType == SplitType.Number)
					{
						splitterRequestModel.SplitNumber = int.Parse( pars);
					}

					AsposeSlides asposeSlides = new AsposeSlides();
					FileSafeResult FileSafeResult = asposeSlides.Splitter(splitterRequestModel, default(System.Threading.CancellationToken));

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

		public ActionResult Splitter()
		{
			var model = new ViewModel(this, "Splitter")
			{
				ControlsView = "SplitterControls",
				SaveAsComponent = true,
				MaximumUploadFiles = 1,
				DropOrUploadFileLabel = Resources["DropOrUploadFile"]
			};
			if (model.RedirectToMainApp)
				return Redirect("/slides/" + model.AppName.ToLower());
			return View(model);
		}
		

	}
}
