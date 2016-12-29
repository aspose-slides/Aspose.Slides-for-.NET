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
    public class ChangeChartCategoryAxis
    {
        public static void Run()
        {
            //ExStart:ChangeChartCategoryAxis
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            using (Presentation presentation = new Presentation(dataDir + "ExistingChart.pptx"))
            {
                IChart chart = presentation.Slides[0].Shapes[0] as IChart;
                chart.Axes.HorizontalAxis.CategoryAxisType = CategoryAxisType.Date;
                chart.Axes.HorizontalAxis.IsAutomaticMajorUnit = false;
                chart.Axes.HorizontalAxis.MajorUnit = 1;
                chart.Axes.HorizontalAxis.MajorUnitScale = TimeUnitType.Months;
                presentation.Save(dataDir + "ChangeChartCategoryAxis_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:ChangeChartCategoryAxis
        }
    }
}