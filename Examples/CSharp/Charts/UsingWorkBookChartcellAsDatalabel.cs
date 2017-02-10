using Aspose.Slides.Charts;
using Aspose.Slides.Export;
using Aspose.Slides.Animation;
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
    public class UsingWorkBookChartcellAsDatalabel
    {
        public static void Run()
        {
            //ExStart:UsingWorkBookChartcellAsDatalabel
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();



            string lbl0 = "Label 0 cell value";
            string lbl1 = "Label 1 cell value";
            string lbl2 = "Label 2 cell value";

            // Instantiate Presentation class that represents a presentation file 

            using (Presentation pres = new Presentation(dataDir + "chart2.pptx"))
            {
                ISlide slide = pres.Slides[0];


                IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.Bubble, 50, 50, 600, 400, true);

                IChartSeriesCollection series = chart.ChartData.Series;

                series[0].Labels.DefaultDataLabelFormat.ShowLabelValueFromCell = true;

                IChartDataWorkbook wb = chart.ChartData.ChartDataWorkbook;

                series[0].Labels[0].ValueFromCell = wb.GetCell(0, "A10", lbl0);
                series[0].Labels[1].ValueFromCell = wb.GetCell(0, "A11", lbl1);
                series[0].Labels[2].ValueFromCell = wb.GetCell(0, "A12", lbl2);

                pres.Save(path + "resultchart.pptx", Aspose.Slides.Export.SaveFormat.Pptx);

            }
       //ExEnd:UsingWorkBookChartcellAsDatalabel
        
        
        }
    }
}