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
    public class SetAutomaticSeriesFillColor
    {
        public static void Run()
        {
            //ExStart:SetAutomaticSeriesFillColor
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();
            using (Presentation presentation = new Presentation())
            {
                // Creating a clustered column chart
                IChart chart = presentation.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 100, 50, 600, 400);

                // Setting series fill format to automatic
                for (int i = 0; i < chart.ChartData.Series.Count; i++)
                {
                    chart.ChartData.Series[i].GetAutomaticSeriesColor();
                }

                // Write the presentation file to disk
                presentation.Save(dataDir + "AutoFillSeries_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:SetAutomaticSeriesFillColor
        }
    }
}