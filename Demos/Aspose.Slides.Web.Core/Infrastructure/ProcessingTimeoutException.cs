using System;

namespace Aspose.Slides.Web.Core.Infrastructure
{
	/// <summary>
	/// Throws at timeout for processing files.
	/// </summary>
	[Serializable]
	public sealed class ProcessingTimeoutException: Exception
	{
		public ProcessingTimeoutException(string message) : base(message)
		{
		}

		public ProcessingTimeoutException(string message, Exception innerException) : base(message, innerException)
		{
		}
	}
}
