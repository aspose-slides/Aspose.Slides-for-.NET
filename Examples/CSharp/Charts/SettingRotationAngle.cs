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
	class SettingRotationAngle
	{
		public static void Run()
		{
			//ExStart:SettingRotationAngle
			// The path to the documents directory.
			string dataDir = RunExamples.GetDataDir_Charts();
			using (Presentation pres = new Presentation())
			{
				IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 450, 300);
				chart.Axes.VerticalAxis.HasTitle = true;
                chart.Axes.VerticalAxis.Title.TextFormat.TextBlockFormat.RotationAngle = 90;

				pres.Save(dataDir + "test.pptx", SaveFormat.Pptx);
			}
		    //ExEnd:SettingRotationAngle

		}
	}
}