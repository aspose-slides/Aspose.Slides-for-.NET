using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Charts
{
    class AddColorToDataPoints
    {
        public static void Run() {

            //ExStart:AddColorToDataPoints

            using (Presentation pres = new Presentation())
            {
                // The path to the documents directory.
                string dataDir = RunExamples.GetDataDir_Charts();


                 IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.Sunburst, 100, 100, 450, 400);

                IChartDataPointCollection dataPoints = chart.ChartData.Series[0].DataPoints;
                dataPoints[3].DataPointLevels[0].Label.DataLabelFormat.ShowValue = true;


            
                IDataLabel branch1Label = dataPoints[0].DataPointLevels[2].Label;
                branch1Label.DataLabelFormat.ShowCategoryName = false;
                branch1Label.DataLabelFormat.ShowSeriesName = true;

                branch1Label.DataLabelFormat.TextFormat.PortionFormat.FillFormat.FillType = FillType.Solid;
                branch1Label.DataLabelFormat.TextFormat.PortionFormat.FillFormat.SolidFillColor.Color = Color.Yellow;


                IFormat steam4Format = dataPoints[9].Format;
                steam4Format.Fill.FillType = FillType.Solid;
               
                steam4Format.Fill.SolidFillColor.Color = Color.FromArgb(0, 176, 240, 255);

                pres.Save(dataDir + "AddColorToDataPoints.pptx", SaveFormat.Pptx);
            }
            //ExEnd:AddColorToDataPoints

        }
    }
}
