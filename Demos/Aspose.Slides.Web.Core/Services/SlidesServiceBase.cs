using Microsoft.Extensions.Logging;

namespace Aspose.Slides.Web.Core.Services
{
	public class SlidesServiceBase
	{
		protected readonly ILogger _logger;

		public SlidesServiceBase(ILogger logger)
		{
			_logger = logger;
		}
	}
}
