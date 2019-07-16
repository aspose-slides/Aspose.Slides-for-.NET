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
    class HideInformationFromChart
    {
        public static void Run() {

            //ExStart:HideInformationFromChart
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            using (Presentation pres = new Presentation())
            {
                ISlide slide = pres.Slides[0];
                IChart chart = slide.Shapes.AddChart(ChartType.LineWithMarkers, 140, 118, 320, 370);

                //Hiding chart Title
                chart.HasTitle = false;

                ///Hiding Values axis
                chart.Axes.VerticalAxis.IsVisible = false;

                //Category Axis visibility
                chart.Axes.HorizontalAxis.IsVisible = false;

                //Hiding Legend
                chart.HasLegend = false;

                //Hiding MajorGridLines
                chart.Axes.HorizontalAxis.MajorGridLinesFormat.Line.FillFormat.FillType = FillType.NoFill;

                for (int i = 0; i < chart.ChartData.Series.Count; i++)
                {
                    chart.ChartData.Series.RemoveAt(i);
                }

                IChartSeries series = chart.ChartData.Series[0];

                series.Marker.Symbol = MarkerStyleType.Circle;
                series.Labels.DefaultDataLabelFormat.ShowValue = true;
                series.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.Top;
                series.Marker.Size = 15;

                //Setting series line color
                series.Format.Line.FillFormat.FillType = FillType.Solid;
                series.Format.Line.FillFormat.SolidFillColor.Color = Color.Purple;
                series.Format.Line.DashStyle = LineDashStyle.Solid;

                pres.Save(dataDir + "HideInformationFromChart.pptx", SaveFormat.Pptx);
            }
            //ExEnd:HideInformationFromChart
        }
    }
}
