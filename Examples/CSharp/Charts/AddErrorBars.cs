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
    public class AddErrorBars
    {
        public static void Run()
        {
            //ExStart:AddErrorBars
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            // Creating empty presentation
            using (Presentation presentation = new Presentation())
            {
                // Creating a bubble chart
                IChart chart = presentation.Slides[0].Shapes.AddChart(ChartType.Bubble, 50, 50, 400, 300, true);

                // Adding Error bars and setting its format
                IErrorBarsFormat errBarX = chart.ChartData.Series[0].ErrorBarsXFormat;
                IErrorBarsFormat errBarY = chart.ChartData.Series[0].ErrorBarsYFormat;
                errBarX.IsVisible = true;
                errBarY.IsVisible = true;
                errBarX.ValueType = ErrorBarValueType.Fixed;
                errBarX.Value = 0.1f;
                errBarY.ValueType = ErrorBarValueType.Percentage;
                errBarY.Value = 5;
                errBarX.Type = ErrorBarType.Plus;
                errBarY.Format.Line.Width = 2;
                errBarX.HasEndCap = true;

                // Saving presentation
                presentation.Save(dataDir + "ErrorBars_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:AddErrorBars
        }
    }
}