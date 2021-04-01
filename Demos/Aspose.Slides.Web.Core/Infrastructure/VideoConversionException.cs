using System;

namespace Aspose.Slides.Web.Core.Infrastructure
{
	/// <summary>
	/// Throws when error video conversion error occurs.
	/// </summary>
	[Serializable]
	internal class VideoConversionException : Exception
	{
		public VideoConversionException(string message) : base(message)
		{
		}

		public VideoConversionException(string message, Exception innerException) : base(message, innerException)
		{
		}
	}
}
