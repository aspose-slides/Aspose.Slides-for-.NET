using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Charts
{
    class SupportOfBubbleSizeRepresentation
    {

        public static void Run() {


            //ExStart:SupportOfBubbleSizeRepresentation
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            using (Presentation pres = new Presentation())
            {
                IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.Bubble, 50, 50, 600, 400, true);

                chart.ChartData.SeriesGroups[0].BubbleSizeRepresentation = BubbleSizeRepresentationType.Width;

                pres.Save(dataDir+ "Presentation_BubbleSizeRepresentation.pptx", SaveFormat.Pptx);
            }

            //ExEnd:SupportOfBubbleSizeRepresentation

        }
    }
}
