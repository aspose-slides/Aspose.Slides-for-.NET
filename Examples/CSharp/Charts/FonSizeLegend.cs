using Aspose.Slides.Charts;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Charts
{

	public class FontSizeLegend
	{
		public static void Run()
		{
			//ExStart:FontSizeLegend
			// The path to the documents directory.
			string dataDir = RunExamples.GetDataDir_Charts();

			using (Presentation pres = new Presentation(dataDir+"test.pptx"))
			{
				IChart chart = pres.Slides[0].Shapes.AddChart(Aspose.Slides.Charts.ChartType.ClusteredColumn, 50, 50, 600, 400);

				chart.Legend.TextFormat.PortionFormat.FontHeight = 20;

				chart.Axes.VerticalAxis.IsAutomaticMinValue = false;

				chart.Axes.VerticalAxis.MinValue = -5;

				chart.Axes.VerticalAxis.IsAutomaticMaxValue = false;

				chart.Axes.VerticalAxis.MaxValue = 10;

				pres.Save(dataDir+"output.pptx", SaveFormat.Pptx);
			}

			//ExEnd:FontSizeLegend
		}
	}
}
	