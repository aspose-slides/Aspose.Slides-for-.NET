using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace Aspose.Slides.Web.Core.Services.Comparison
{
	/// <summary>
	/// The implementation for strategy of compare text by slides
	/// </summary>
	public sealed class SlidesTextComparator : BaseComparator, IPresentationComparator
	{
		/// <summary>
		/// Compares two presentations and returns a string with differents.
		/// </summary>
		/// <param name="firstPresentationFile">A string path of the first presentation file</param>
		/// <param name="secondPresentationFile">A string path of the second presentation file</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns>String with differences</returns>
		public override string ComparePresentations(string firstPresentationFile, string secondPresentationFile, CancellationToken cancellationToken = default)
		{
			using var firstPresentation = new Presentation(firstPresentationFile);
			using var secondPresentation = new Presentation(secondPresentationFile);

			IPresentation leftPresentation;
			IPresentation rightPresentation;
			string leftPresentationName;
			string rightPresentationName;

			if (firstPresentation.Slides.Count >= secondPresentation.Slides.Count)
			{
				leftPresentation = firstPresentation;
				leftPresentationName = Path.GetFileName(firstPresentationFile);
				rightPresentation = secondPresentation;
				rightPresentationName = Path.GetFileName(secondPresentationFile);
			}
			else // firstPresentation.Slides.Count < secondPresentation.Slides.Count
			{
				leftPresentation = secondPresentation;
				leftPresentationName = Path.GetFileName(secondPresentationFile);
				rightPresentation = firstPresentation;
				rightPresentationName = Path.GetFileName(firstPresentationFile);
			}

			cancellationToken.ThrowIfCancellationRequested();

			var diffTextSummary = new StringBuilder();

			foreach (var leftSlide in leftPresentation.Slides)
			{
				cancellationToken.ThrowIfCancellationRequested();

				var slideIndex = leftPresentation.Slides.IndexOf(leftSlide);
				ISlide rightSlide = null;

				if (slideIndex < rightPresentation.Slides.Count)
				{
					rightSlide = rightPresentation.Slides[slideIndex];
				}

				var diffText = CompareSlides(leftSlide, rightSlide, leftPresentationName, rightPresentationName, slideIndex);

				if (!String.IsNullOrWhiteSpace(diffText))
				{
					diffTextSummary.Append(diffText);
				}

				cancellationToken.ThrowIfCancellationRequested();
			}

			return diffTextSummary.ToString();
		}

		private string CompareSlides(ISlide leftSlide, ISlide rightSlide, string leftPresentationName, string rightPresentationName, int slideIndex)
		{
			var diffTextSummary = new StringBuilder();
			var slideNumber = slideIndex + 1;

			IShapeCollection leftShapes;
			IShapeCollection rightShapes;

			if (rightSlide != null)
			{
				if (leftSlide.Shapes.Count >= rightSlide.Shapes.Count)
				{
					leftShapes = leftSlide.Shapes;
					rightShapes = rightSlide.Shapes;
				}
				else
				{
					leftShapes = rightSlide.Shapes;
					rightShapes = leftSlide.Shapes;

					var tempName = leftPresentationName;
					leftPresentationName = rightPresentationName;
					rightPresentationName = tempName;
				}
			}
			else
			{
				leftShapes = leftSlide.Shapes;
				rightShapes = null;
			}

			var sortedLeftShapes = leftShapes.OrderBy((shape) => shape.Frame.Y).Where((shape) => shape is AutoShape).Cast<AutoShape>().ToList();
			var sortedRightShapes = rightShapes == null ? null : rightShapes.OrderBy((shape) => shape.Frame.Y).Where((shape) => shape is AutoShape).Cast<AutoShape>().ToList();					

			foreach (var leftShape in sortedLeftShapes)
			{
				var index = sortedLeftShapes.IndexOf(leftShape);

				if (sortedRightShapes != null && index < sortedRightShapes.Count)
				{
					var rightShape = sortedRightShapes[index];
				
					if (leftShape.TextFrame.Text != rightShape.TextFrame.Text)
					{
						var diffText = base.PrepareDiffText(leftShape.TextFrame.Text, rightShape.TextFrame.Text, leftPresentationName, rightPresentationName, slideNumber );

						diffTextSummary.Append(diffText);
					}
				}
				else
				{
					var diffText = base.PrepareDiffText(leftShape.TextFrame.Text, String.Empty, leftPresentationName, rightPresentationName, slideNumber);

					diffTextSummary.Append(diffText);
				}
			}

			return diffTextSummary.ToString();
		}
	}
}
