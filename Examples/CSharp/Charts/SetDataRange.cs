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
    public class SetDataRange
    {
        public static void Run()
        {
            //ExStart:SetDataRange
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            // Instantiate Presentation class that represents PPTX file
            Presentation presentation = new Presentation(dataDir + "ExistingChart.pptx");

            // Access first slideMarker and add chart with default data
            ISlide slide = presentation.Slides[0];
            IChart chart = (IChart)slide.Shapes[0];
            chart.ChartData.SetRange("Sheet1!A1:B4");
            presentation.Save(dataDir + "SetDataRange_out.pptx", SaveFormat.Pptx);
            //ExEnd:SetDataRange
        }
    }
}