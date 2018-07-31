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
	class SupportForChartAreaRoundedBorders
	{
		public static void Run()
		{
			//ExStart:SupportForChartAreaRoundedBorders
			string dataDir = RunExamples.GetDataDir_Charts();
			using (Presentation presentation = new Presentation())
			{
				ISlide slide = presentation.Slides[0];
				IChart chart = slide.Shapes.AddChart(ChartType.ClusteredColumn, 20, 100, 600, 400);
				chart.LineFormat.FillFormat.FillType = FillType.Solid;
				chart.LineFormat.Style = LineStyle.Single;
				chart.HasRoundedCorners = true;

				presentation.Save(dataDir + "out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
			}
		}	
		//ExEnd:SupportForChartAreaRoundedBorders
	}
}
