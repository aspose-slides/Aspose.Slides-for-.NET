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
    class FontPropertiesForChart
    {
        public static void Run() {

            //ExStart:FontPropertiesForChart
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            using (Presentation pres = new Presentation())
            {               

                IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 500, 400);
                chart.TextFormat.PortionFormat.FontHeight = 20;
                chart.ChartData.Series[0].Labels.DefaultDataLabelFormat.ShowValue = true;
                pres.Save(dataDir + "FontPropertiesForChart.pptx", SaveFormat.Pptx);
            }

            //ExEnd:FontPropertiesForChart

        }
    }
}
