using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Charts
{
	class SupportForChangingSeriesColor
	{
		public static void Run()
		{
			//ExStart:SupportForChangingSeriesColor
			// The path to the documents directory.
			string dataDir = RunExamples.GetDataDir_Charts();

			using (Presentation pres = new Presentation(dataDir+"test.pptx"))
			{
				IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.Pie, 50, 50, 600, 400);

				IChartDataPoint point = chart.ChartData.Series[0].DataPoints[1];

				point.Explosion = 30;

				point.Format.Fill.FillType = FillType.Solid;

				point.Format.Fill.SolidFillColor.Color = Color.Blue;


				pres.Save(dataDir+"output.pptx", SaveFormat.Pptx);
			}
			//ExEnd:SupportForChangingSeriesColor

		}
	}
}