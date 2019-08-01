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
    class DefaultMarkersInChart
    {
        public static void Run() {

            //ExStart:DefaultMarkersInChart
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();


            using (Presentation pres = new Presentation())
            {
                ISlide slide = pres.Slides[0];
                IChart chart = slide.Shapes.AddChart(ChartType.LineWithMarkers, 10, 10, 400, 400);

                chart.ChartData.Series.Clear();
                chart.ChartData.Categories.Clear();

                IChartDataWorkbook fact = chart.ChartData.ChartDataWorkbook;
                chart.ChartData.Series.Add(fact.GetCell(0, 0, 1, "Series 1"), chart.Type);
                IChartSeries series = chart.ChartData.Series[0];

                chart.ChartData.Categories.Add(fact.GetCell(0, 1, 0, "C1"));
                series.DataPoints.AddDataPointForLineSeries(fact.GetCell(0, 1, 1, 24));
                chart.ChartData.Categories.Add(fact.GetCell(0, 2, 0, "C2"));
                series.DataPoints.AddDataPointForLineSeries(fact.GetCell(0, 2, 1, 23));
                chart.ChartData.Categories.Add(fact.GetCell(0, 3, 0, "C3"));
                series.DataPoints.AddDataPointForLineSeries(fact.GetCell(0, 3, 1, -10));
                chart.ChartData.Categories.Add(fact.GetCell(0, 4, 0, "C4"));
                series.DataPoints.AddDataPointForLineSeries(fact.GetCell(0, 4, 1, null));

                chart.ChartData.Series.Add(fact.GetCell(0, 0, 2, "Series 2"), chart.Type);
                //Take second chart series
                IChartSeries series2 = chart.ChartData.Series[1];

                //Now populating series data
                series2.DataPoints.AddDataPointForLineSeries(fact.GetCell(0, 1, 2, 30));
                series2.DataPoints.AddDataPointForLineSeries(fact.GetCell(0, 2, 2, 10));
                series2.DataPoints.AddDataPointForLineSeries(fact.GetCell(0, 3, 2, 60));
                series2.DataPoints.AddDataPointForLineSeries(fact.GetCell(0, 4, 2, 40));

                chart.HasLegend = true;
                chart.Legend.Overlay = false;

                pres.Save(dataDir + "DefaultMarkersInChart.pptx", SaveFormat.Pptx);
            }

            //ExEnd:DefaultMarkersInChart

        }
    }
}
