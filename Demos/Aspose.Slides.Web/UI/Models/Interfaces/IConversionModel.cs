namespace Aspose.Slides.Web.UI.Models.Interfaces
{
	public interface IConversionModel : IBaseViewModel
	{
		string FromFormat { get; }
		string ToFormat { get; }
		string[] Formats { get; }
	}
}
