using Aspose.Slides.Live.Demos.UI.Models.Common;
using Aspose.Slides.Live.Demos.UI.Models;
using Aspose.Slides.Live.Demos.UI.Services;
using System;
using System.Collections;
using System.Web;
using System.Web.Mvc;
using System.Threading.Tasks;

namespace Aspose.Slides.Live.Demos.UI.Controllers
{
	public class ConversionController : BaseController  
	{
		public override string Product => (string)RouteData.Values["product"];


		[HttpPost]
		public  Response Conversion(string outputType)
		{
			Response response = null;
			if (Request.Files.Count > 0)
			{
				string _sourceFolder = Guid.NewGuid().ToString();
				var docs = UploadFiles(Request, _sourceFolder);

				if (docs.Count > 0)
				{
					ConversionOptions conversionOptions = new ConversionOptions();
					conversionOptions.id = docs[0].FolderName;
					conversionOptions.FileName = docs[0].FileName;
					conversionOptions.Format = outputType.Trim();

					AsposeSlides asposeSlides = new AsposeSlides();
					FileSafeResult FileSafeResult =  asposeSlides.Conversion(conversionOptions);

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

		

		public ActionResult Conversion()
		{
			var model = new ViewModel(this, "Conversion")
			{
				SaveAsComponent = true,
				SaveAsOriginal = false,
				MaximumUploadFiles = 1,
				DropOrUploadFileLabel = Resources["DropOrUploadFile"]
			};

			return View(model);
		}
		

	}
}
