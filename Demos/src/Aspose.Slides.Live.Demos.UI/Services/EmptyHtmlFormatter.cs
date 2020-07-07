using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Live.Demos.UI.Services
{
	internal class EmptyHtmlFormatter : IHtmlFormattingController
	{
		public void WriteDocumentStart(IHtmlGenerator generator, IPresentation presentation)
		{
		}

		public void WriteDocumentEnd(IHtmlGenerator generator, IPresentation presentation)
		{
		}

		public void WriteSlideStart(IHtmlGenerator generator, ISlide slide)
		{
		}

		public void WriteSlideEnd(IHtmlGenerator generator, ISlide slide)
		{
		}

		public void WriteShapeStart(IHtmlGenerator generator, IShape shape)
		{
		}

		public void WriteShapeEnd(IHtmlGenerator generator, IShape shape)
		{
		}
	}
}
