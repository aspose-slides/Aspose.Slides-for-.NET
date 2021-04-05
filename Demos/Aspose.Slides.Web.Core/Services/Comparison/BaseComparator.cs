using System;
using System.Text;
using System.Threading;

namespace Aspose.Slides.Web.Core.Services.Comparison
{
	/// <summary>
	/// The abstract base class for strategy of compare
	/// </summary>
	public abstract class BaseComparator : IPresentationComparator
	{
		/// <summary>
		/// Compares two presentations and returns a string with differents.
		/// </summary>
		/// <param name="firstPresentationFile">A string path of the first presentation file</param>
		/// <param name="secondPresentationFile">A string path of the second presentation file</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns>String with differences</returns>
		public abstract string ComparePresentations(string firstPresentationFile, string secondPresentationFile, CancellationToken cancellationToken = default);

		private const string DiffTextLimiter = "<<<<<<<<<<";
		private const string DiffTextSeparator = "==========";

		/// <summary>
		/// Prepares diff text for single diff
		/// </summary>
		/// <param name="leftSlideDiffText"></param>
		/// <param name="rightSlideDiffText"></param>
		/// <param name="leftPresentationName"></param>
		/// <param name="rightPresentationName"></param>
		/// <param name="slideNumber"></param>
		/// <returns></returns>
		protected string PrepareDiffText(string leftSlideDiffText, string rightSlideDiffText, string leftPresentationName, string rightPresentationName, int slideNumber)
		{
			var diffTextSummary = new StringBuilder();

			diffTextSummary.AppendLine(DiffTextLimiter);
			diffTextSummary.AppendLine($"{leftPresentationName} slide {slideNumber}:");
			diffTextSummary.AppendLine(leftSlideDiffText);
			diffTextSummary.AppendLine(DiffTextSeparator);
			diffTextSummary.AppendLine($"{rightPresentationName} slide {slideNumber}:");
			diffTextSummary.AppendLine(rightSlideDiffText);
			diffTextSummary.AppendLine(DiffTextLimiter);
			diffTextSummary.AppendLine(Environment.NewLine);

			return diffTextSummary.ToString();
		}
	}
}
