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
    class ClearSpecificChartSeriesDataPointsData
    {
        public static void Run() {

            //ExStart:ClearSpecificChartSeriesDataPointsData

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            using (Presentation pres = new Presentation(dataDir + "TestChart.pptx"))
            {
                ISlide sl = pres.Slides[0];

                IChart chart = (IChart)sl.Shapes[0];

                foreach (IChartDataPoint dataPoint in chart.ChartData.Series[0].DataPoints)
                {
                    dataPoint.XValue.AsCell.Value = null;
                    dataPoint.YValue.AsCell.Value = null;
                }

                chart.ChartData.Series[0].DataPoints.Clear();

                pres.Save(dataDir + "ClearSpecificChartSeriesDataPointsData.pptx", SaveFormat.Pptx);
            }

            //ExEnd:ClearSpecificChartSeriesDataPointsData

        }
    }
}
