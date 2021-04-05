namespace Aspose.Slides.Web.Interfaces.Services
{
	public interface ITemporaryStorage
	{
		ITemporaryFolder GetTemporaryFolder();

		ITemporaryFolder GetTemporaryFolder(string id);
	}
}
