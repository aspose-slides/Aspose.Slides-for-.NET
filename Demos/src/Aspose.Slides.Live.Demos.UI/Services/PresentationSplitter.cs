using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Slides.Live.Demos.UI.Models;
using Aspose.Slides;

namespace Aspose.Slides.Live.Demos.UI.Services
{
	internal static class PresentationSplitter
	{
		public static IEnumerable<(string name, ISlide[] slides)> GetChunks(this Presentation presentation, SplitType splitType, int splitNumber, string splitRange)
		{
			switch (splitType)
			{
				case SplitType.SlideBySlide:
					return presentation.Slides.Select(s => (s.SlideNumber.ToString("D2"), new[] { s })).ToArray();
				case SplitType.EvenOdd:
					return new[]
					{
						("odd", presentation.Slides.Where(s => s.SlideNumber % 2 != 0).ToArray()),
						("even", presentation.Slides.Where(s => s.SlideNumber % 2 == 0).ToArray()),
					};
				case SplitType.Number:
					if (splitNumber <= 0)
					{
						return new (string name, ISlide[] slides)[]{};
					}

					return presentation.Slides
						.GroupBy(s => (s.SlideNumber - 1) / splitNumber)
						.Select(g => ((g.Key + 1).ToString("D2"), g.ToArray())).ToArray();
				case SplitType.Range:
					var ranges = splitRange.Split(',').Select(range =>
					{
						var bounds = range.Split('-');
						var start = Convert.ToInt32(bounds[0].Trim());
						if (start < 0)
						{
							return new int[] { };
						}

						var count = bounds.Length > 1 ? Convert.ToInt32(bounds[1].Trim()) - start + 1 : 1;
						if (count < 0)
						{
							count = 0;
						}
						return Enumerable.Range(start, count);
					});
					return ranges.Select((r, i) => ((i + 1).ToString("D2"),
							r.Select(slideNumber => slideNumber < 1 || slideNumber > presentation.Slides.Count ? null :
									presentation.Slides[slideNumber - 1])
								.Where(s => s != null)
								.ToArray()))
						.ToArray();
				default:
					throw new ArgumentOutOfRangeException(nameof(splitType), splitType, null);
			}
		}

	}
}
