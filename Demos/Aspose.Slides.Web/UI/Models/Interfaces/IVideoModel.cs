namespace Aspose.Slides.Web.UI.Models.Interfaces
{
	public interface IVideoModel : IBaseViewModel
	{
		public string Every { get; }
		public string PageRange { get; }
		public string RangeException { get; }

		string[] VideoCodecs { get; }
	}
}
