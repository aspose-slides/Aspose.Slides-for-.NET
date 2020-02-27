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
			using (Presentation pres = new Presentation(dataDir + "Test.pptx"))
			{
				IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 400, 300);

			    IChartSeries[] series = new IChartSeries[chart.ChartData.Series.Count];
			    chart.ChartData.Series.CopyTo(series, 0);

			    IChartDataCell[] categoriesCells = new IChartDataCell[chart.ChartData.Categories.Count];

			    for (int i = 0; i < chart.ChartData.Categories.Count; i++)
			    {
			        categoriesCells[i] = chart.ChartData.Categories[i].AsCell;
			    }

			    IChartDataCell[] seriesCells = new IChartDataCell[chart.ChartData.Series.Count];
			    for (int i = 0; i < chart.ChartData.Series.Count; i++)
			    {
			        seriesCells[i] = chart.ChartData.Series[i].Name.AsCells[0];
			    }

			    chart.ChartData.SwitchRowColumn();

			    pres.Save(RunExamples.OutPath + "Test_out.pptx", SaveFormat.Pptx);
				//ExEnd:SupportForSwitchingRowsAndColumns
			}

		}
	}
}