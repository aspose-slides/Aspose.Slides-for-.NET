using System.Net;
using System.Threading.Tasks;
using Aspose.Slides.Web.UI.Services;
using Microsoft.AspNetCore.Mvc;

namespace Aspose.Slides.Web.UI.Controllers
{
	public class SlidesController : Controller
	{
		public SlidesController(ISlidesViewModelFactory slidesViewModelFactory, IEditorService editorService)
		{
			SlidesViewModelFactory = slidesViewModelFactory;
			EditorService = editorService;
		}

		public ISlidesViewModelFactory SlidesViewModelFactory { get; }
		public IEditorService EditorService { get; }

		public ActionResult Index()
		{
			return View(SlidesViewModelFactory.CreateShowroomModel(Request));
		}

		public ActionResult Annotation(string extension)
		{
			var model = SlidesViewModelFactory.CreateAnnotationModel(Request, extension);
			return View(model);
		}

		public ActionResult Search(string extension)
		{
			var model = SlidesViewModelFactory.CreateSearchModel(Request, extension);

			return View(model);
		}

		public ActionResult Watermark(string extension)
		{
			var model = SlidesViewModelFactory.CreateWatermarkModel(Request, extension);

			return View(model);
		}

		public ActionResult Redaction(string extension)
		{
			var model = SlidesViewModelFactory.CreateRedactionModel(Request, extension);

			return View(model);
		}

		public ActionResult Parser(string extension)
		{
			var model = SlidesViewModelFactory.CreateParserModel(Request, extension);

			return View(model);
		}

		public ActionResult Conversion(string extension)
		{
			var model = SlidesViewModelFactory.CreateConversionModel(Request, extension);

			return View(model);
		}

		public ActionResult Merger(string extension)
		{
			var model = SlidesViewModelFactory.CreateMergerModel(Request, extension);

			return View(model);
		}

		[Route("slides/storage/view/{folder?}/{fileName?}")]
		public ActionResult Slideshow(string folder, string fileName)
		{
			if (string.IsNullOrEmpty(folder) || string.IsNullOrEmpty(fileName))
			{
				return Redirect("/slides/viewer");
			}

			var model = SlidesViewModelFactory.CreateSlideshowModel(Request, folder, fileName);

			return View(model);
		}

		public ActionResult Viewer(string extension)
		{
			var model = SlidesViewModelFactory.CreateViewerModel(Request, extension);

			return View(model);
		}

		public ActionResult Unlock(string extension)
		{
			var model = SlidesViewModelFactory.CreateUnlockModel(Request, extension);

			return View(model);
		}

		public ActionResult Lock(string extension)
		{
			var model = SlidesViewModelFactory.CreateLockModel(Request, extension);

			return View(model);
		}

		public ActionResult Metadata(string extension)
		{
			var model = SlidesViewModelFactory.CreateMetadataModel(Request, extension);

			return View(model);
		}

		public ActionResult Video(string extension)
		{
			var model = SlidesViewModelFactory.CreateVideoModel(Request, extension);

			return View(model);
		}

		public ActionResult Splitter(string extension)
		{
			var model = SlidesViewModelFactory.CreateSplitterModel(Request, extension);

			return View(model);
		}

		[Route("slides/storage/edit/{copy?}/{new?}/{folder?}/{fileName?}")]
		public async Task<ActionResult> EditorApp(string copy, string @new, string folder, string fileName)
		{
			if (!string.IsNullOrEmpty(copy))
			{
				var res = await EditorService.CopyProcessedAsync(folder, fileName);

				if (res == null)
				{
					return Redirect("/slides/storage/edit/?folder=NotFound&fileName=NotFound");
				}

				return Redirect(
					$"/slides/storage/edit/?folder={res.Folder}&fileName={WebUtility.UrlEncode(res.Filename)}"
				);
			}

			if (!string.IsNullOrEmpty(@new))
			{
				var res = await EditorService.CreateByTemplateAsync(@new);

				if (res == null)
				{
					return Redirect("/slides/storage/edit/?folder=NotFound&fileName=NotFound");
				}

				return Redirect(
					$"/slides/storage/edit/?folder={res.Folder}&fileName={WebUtility.UrlEncode(res.Filename)}"
				);
			}

			var model = SlidesViewModelFactory.CreateEditorAppModel(Request, folder, fileName);

			return View(model);
		}

		public ActionResult Editor(string extension)
		{
			var model = SlidesViewModelFactory.CreateEditorUploaderModel(Request, extension);

			return View(model);
		}

		public ActionResult Signature(string extension)
		{
			var model = SlidesViewModelFactory.CreateSignatureModel(Request, extension);

			return View(model);
		}

		public ActionResult Chart(string extension)
		{
			var model = SlidesViewModelFactory.CreateChartModel(Request, extension);

			return View(model);
		}

		public ActionResult PermanentlyRedirect(string toaction, string extension)
		{
			return RedirectPermanent($"/slides/{toaction}/{extension}");
		}

		public ActionResult Comparison(string extension)
		{
			var model = SlidesViewModelFactory.CreateComparisonModel(Request, extension);

			return View(model);
		}

		public ActionResult Import(string extension)
		{
			var model = SlidesViewModelFactory.CreateImportModel(Request, extension);

			return View(model);
		}

		[ActionName("Remove-Macros")]
		public ActionResult RemoveMacros(string extension)
		{
			var model = SlidesViewModelFactory.CreateRemoveMacrosModel(Request, extension);
			return View("RemoveMacros", model);
		}
	}
}
