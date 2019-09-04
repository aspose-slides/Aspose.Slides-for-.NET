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
    class SetLayoutMode
    {
        public static void Run() {

            //ExStart:SetLayoutMode
            string dataDir = RunExamples.GetDataDir_Charts();
            using (Presentation presentation = new Presentation())
            {
                ISlide slide = presentation.Slides[0];
                IChart chart = slide.Shapes.AddChart(ChartType.ClusteredColumn, 20, 100, 600, 400);
                chart.PlotArea.AsILayoutable.X = 0.2f;
                chart.PlotArea.AsILayoutable.Y = 0.2f;
                chart.PlotArea.AsILayoutable.Width = 0.7f;
                chart.PlotArea.AsILayoutable.Height = 0.7f;
                chart.PlotArea.LayoutTargetType = LayoutTargetType.Inner;

                presentation.Save(dataDir + "SetLayoutMode_outer.pptx", SaveFormat.Pptx);
   
}
            //ExEnd:SetLayoutMode
        }
    }
}
