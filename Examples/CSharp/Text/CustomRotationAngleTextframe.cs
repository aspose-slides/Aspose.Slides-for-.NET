using Aspose.Slides.Export;
using Aspose.Slides.Charts;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Text
{
    class CustomRotationAngleTextframe
    {
        public static void Run()
        {
            // ExStart:CustomRotationAngleTextframe

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            // ExStart:CustomRotationAngleTextframe
            // Create an instance of Presentation class
            Presentation presentation = new Presentation();

            IChart chart = presentation.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 500, 300);

            IChartSeries series = chart.ChartData.Series[0];

            series.Labels.DefaultDataLabelFormat.ShowValue = true;
            series.Labels.DefaultDataLabelFormat.TextFormat.TextBlockFormat.RotationAngle = 65;

            chart.HasTitle = true;
            chart.ChartTitle.AddTextFrameForOverriding("Custom title").TextFrameFormat.RotationAngle = -30;

            // ExEnd:CustomRotationAngleTextframe
            // Save Presentation
            presentation.Save(dataDir + "textframe-rotation_out.pptx", SaveFormat.Pptx);
            // ExEnd:CustomRotationAngleTextframe
        }
    }
}