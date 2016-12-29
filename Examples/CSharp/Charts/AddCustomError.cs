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
    public class AddCustomError
    {
        public static void Run()
        {
            //ExStart:AddCustomError
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            // Creating empty presentation
            using (Presentation presentation = new Presentation())
            {
                // Creating a bubble chart
                IChart chart = presentation.Slides[0].Shapes.AddChart(ChartType.Bubble, 50, 50, 400, 300, true);

                // Adding custom Error bars and setting its format
                IChartSeries series = chart.ChartData.Series[0];
                IErrorBarsFormat errBarX = series.ErrorBarsXFormat;
                IErrorBarsFormat errBarY = series.ErrorBarsYFormat;
                errBarX.IsVisible = true;
                errBarY.IsVisible = true;
                errBarX.ValueType = ErrorBarValueType.Custom;
                errBarY.ValueType = ErrorBarValueType.Custom;

                // Accessing chart series data point and setting error bars values for individual point
                IChartDataPointCollection points = series.DataPoints;
                points.DataSourceTypeForErrorBarsCustomValues.DataSourceTypeForXPlusValues = DataSourceType.DoubleLiterals;
                points.DataSourceTypeForErrorBarsCustomValues.DataSourceTypeForXMinusValues = DataSourceType.DoubleLiterals;
                points.DataSourceTypeForErrorBarsCustomValues.DataSourceTypeForYPlusValues = DataSourceType.DoubleLiterals;
                points.DataSourceTypeForErrorBarsCustomValues.DataSourceTypeForYMinusValues = DataSourceType.DoubleLiterals;

                // Setting error bars for chart series points
                for (int i = 0; i < points.Count; i++)
                {
                    points[i].ErrorBarsCustomValues.XMinus.AsLiteralDouble = i + 1;
                    points[i].ErrorBarsCustomValues.XPlus.AsLiteralDouble = i + 1;
                    points[i].ErrorBarsCustomValues.YMinus.AsLiteralDouble = i + 1;
                    points[i].ErrorBarsCustomValues.YPlus.AsLiteralDouble = i + 1;
                }

                // Saving presentation
                presentation.Save(dataDir + "ErrorBarsCustomValues_out.pptx", SaveFormat.Pptx);
                //ExEnd:AddCustomError
            }
        }
    }
}