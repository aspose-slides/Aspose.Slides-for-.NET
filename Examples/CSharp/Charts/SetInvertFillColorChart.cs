using System.Drawing;
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
    public class SetInvertFillColorChart
    {
        public static void Run()
        {
            //ExStart:SetInvertFillColorChart
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();
            Color inverColor = Color.Red;
            using (Presentation pres = new Presentation())
            {
                IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 400, 300);
                IChartDataWorkbook workBook = chart.ChartData.ChartDataWorkbook;

                chart.ChartData.Series.Clear();
                chart.ChartData.Categories.Clear();

                // Adding new series and categories
                chart.ChartData.Series.Add(workBook.GetCell(0, 0, 1, "Series 1"), chart.Type);
                chart.ChartData.Categories.Add(workBook.GetCell(0, 1, 0, "Category 1"));
                chart.ChartData.Categories.Add(workBook.GetCell(0, 2, 0, "Category 2"));
                chart.ChartData.Categories.Add(workBook.GetCell(0, 3, 0, "Category 3"));

                // Take first chart series and populating series data.
                IChartSeries series = chart.ChartData.Series[0];
                series.DataPoints.AddDataPointForBarSeries(workBook.GetCell(0, 1, 1, -20));
                series.DataPoints.AddDataPointForBarSeries(workBook.GetCell(0, 2, 1, 50));
                series.DataPoints.AddDataPointForBarSeries(workBook.GetCell(0, 3, 1, -30));
                var seriesColor = series.GetAutomaticSeriesColor();
                series.InvertIfNegative = true;
                series.Format.Fill.FillType = FillType.Solid;
                series.Format.Fill.SolidFillColor.Color = seriesColor;
                series.InvertedSolidFillColor.Color = inverColor;
                pres.Save(dataDir + "SetInvertFillColorChart_out.pptx", SaveFormat.Pptx);               
            }
            //ExEnd:SetInvertFillColorChart
        }
    }
}