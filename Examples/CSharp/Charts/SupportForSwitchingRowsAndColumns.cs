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
	class SupportForSwitchingRowsAndColumns
	{
        public static void Run()
		{
			//ExStart:SupportForSwitchingRowsAndColumns

			string dataDir = RunExamples.GetDataDir_Charts();
			using (Presentation pres = new Presentation(dataDir+"Test.pptx"))
			{
				IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 400, 300);
				chart.ChartData.SwitchRowColumn();
				pres.Save(dataDir, SaveFormat.Pptx);
				//ExEnd:SupportForSwitchingRowsAndColumns
			}

		}
	}
}