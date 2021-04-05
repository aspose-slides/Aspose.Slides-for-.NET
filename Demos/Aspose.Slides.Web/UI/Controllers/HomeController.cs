using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;

namespace Aspose.Slides.Web.UI.Controllers
{
	public class HomeController : Controller
	{
		private readonly ILogger<HomeController> _logger;

		public HomeController(ILogger<HomeController> logger)
		{
			_logger = logger;
		}

		public IActionResult Index() => RedirectToAction("Index", "Slides");
	}
}
