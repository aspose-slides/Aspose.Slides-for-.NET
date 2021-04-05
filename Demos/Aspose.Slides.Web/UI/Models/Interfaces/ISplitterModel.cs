namespace Aspose.Slides.Web.UI.Models.Interfaces
{
	public interface ISplitterModel : IBaseViewModel
	{
		public string IntoMany { get; }
		public string Every { get; }
		public string Odd { get; }
		public string ByNumber { get; }
		public string IntoSingle { get; }
		public string PageRange { get; }
		public string ToFormat { get; }
		public string[] Formats { get; }
		public string RangeException { get; }
		public string NumberException { get; }
	}
}
