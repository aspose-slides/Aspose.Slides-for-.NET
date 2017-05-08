using Aspose.Slides.Charts;
using Aspose.Slides.Export;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Charts
{
    public class GetActualPositionOFChartDatalabel
    {
        public static void Run()
        {
            //ExStart:GetActualPositionOFChartDatalabel
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();
            using (Presentation pres = new Presentation())
            {
                IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 500, 400);
                foreach (IChartSeries series in chart.ChartData.Series)
                {
                    series.Labels.DefaultDataLabelFormat.Position = LegendDataLabelPosition.OutsideEnd;
                    series.Labels.DefaultDataLabelFormat.ShowValue = true;
                }

                chart.ValidateChartLayout();

                foreach (IChartSeries series in chart.ChartData.Series)
                {
                    foreach (IChartDataPoint point in series.DataPoints)
                    {
                        if (point.Value.ToDouble() > 4)
                        {
                            float x = point.Label.ActualX;
                            float y = point.Label.ActualY;
                            float w = point.Label.ActualWidth;
                            float h = point.Label.ActualHeight;

                            IAutoShape shape = chart.UserShapes.Shapes.AddAutoShape(ShapeType.Ellipse, x, y, w, h);
                            shape.FillFormat.FillType = FillType.Solid;
                            shape.FillFormat.SolidFillColor.Color = Color.FromArgb(100, 0, 255, 0);
                        }
                    }
                }

                pres.Save(pptxFileName, SaveFormat.Pptx);
            }

            //ExEnd:GetActualPositionOFChartDatalabel
        }
    }
}