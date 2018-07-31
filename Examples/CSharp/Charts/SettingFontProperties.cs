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
	class SettingFontProperties
	{
		public static void Run()
		{
			string dataDir = RunExamples.GetDataDir_Charts();
			using (Presentation pres = new Presentation(dataDir+"test.pptx"))
			{

				IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 600, 400);

				chart.HasDataTable = true;

				chart.ChartDataTable.TextFormat.PortionFormat.FontBold = NullableBool.True;
				chart.ChartDataTable.TextFormat.PortionFormat.FontHeight = 20;

				pres.Save(dataDir+"output.pptx", SaveFormat.Pptx);

			}
		}
	}
}