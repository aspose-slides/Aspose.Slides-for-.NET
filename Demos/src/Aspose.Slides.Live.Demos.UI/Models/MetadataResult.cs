namespace Aspose.Slides.Live.Demos.UI.Models
{
	public class MetadataResult : FileSafeResult
	{
		public MetadataResult()
		{

		}

		internal MetadataResult(string localFilePath)
			: base(localFilePath)
		{

		}

		public PresentationMetadata Metadata { get; set; }
	}
}
