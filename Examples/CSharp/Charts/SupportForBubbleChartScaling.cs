using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Charts
{
	class SupportForBubbleChartScaling
	{
		public static void Run()
		{
			//ExStart:SupportForBubbleChartScaling
			string dataDir = RunExamples.GetDataDir_Charts();
			using (Presentation pres = new Presentation())
			{
				IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.Bubble, 100, 100, 400, 300);

				chart.ChartData.SeriesGroups[0].BubbleSizeScale = 150;

				pres.Save(dataDir+"Result.pptx",Aspose.Slides.Export.SaveFormat.Pptx);
			}

			//ExEnd:SupportForBubbleChartScaling

		}
	}
}