using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Charts
{
	class SettingPositionAxis
	{
		public static void Run()
		{
			//ExStart:SettingPositionAxis
			string dataDir = RunExamples.GetDataDir_Charts();
			using (Presentation pres = new Presentation())
			{
				IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 450, 300);
				chart.Axes.HorizontalAxis.AxisBetweenCategories = true;

				pres.Save(dataDir + "AsposeScatterChart.pptx", SaveFormat.Pptx);

			}
            //ExEnd:SettingPositionAxis

        }
    }
}