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
	class ShowingDisplayUnitLabel
	{
		public static void Run()
		{
			//ExStart:ShowingDisplayUnitLabel
			string dataDir = RunExamples.GetDataDir_Charts();
			using (Presentation pres = new Presentation(dataDir+"Test.pptx"))
			{
				IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 450, 300);
				chart.Axes.VerticalAxis.DisplayUnit = DisplayUnitType.Millions;
				pres.Save(dataDir + "Result.pptx", SaveFormat.Pptx);

			}
            //ExEnd:ShowingDisplayUnitLabel
		}
	}
}