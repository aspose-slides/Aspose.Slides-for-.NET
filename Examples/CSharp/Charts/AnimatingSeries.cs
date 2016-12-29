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
    public class AnimatingSeries
    {
        public static void Run()
        {
            //ExStart:AnimatingSeries
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            // Instantiate Presentation class that represents a presentation file 
            using (Presentation presentation = new Presentation(dataDir + "ExistingChart.pptx"))
            {
                // Get reference of the chart object
                var slide = presentation.Slides[0] as Slide;
                var shapes = slide.Shapes as ShapeCollection;
                var chart = shapes[0] as IChart;

                // Animate the series
                slide.Timeline.MainSequence.AddEffect(chart, EffectType.Fade, EffectSubtype.None,
                EffectTriggerType.AfterPrevious);

                ((Sequence)slide.Timeline.MainSequence).AddEffect(chart,
                EffectChartMajorGroupingType.BySeries, 0,
                EffectType.Appear, EffectSubtype.None, EffectTriggerType.AfterPrevious);

                ((Sequence)slide.Timeline.MainSequence).AddEffect(chart,
                EffectChartMajorGroupingType.BySeries, 1,
                EffectType.Appear, EffectSubtype.None, EffectTriggerType.AfterPrevious);

                ((Sequence)slide.Timeline.MainSequence).AddEffect(chart,
                EffectChartMajorGroupingType.BySeries, 2,
                EffectType.Appear, EffectSubtype.None, EffectTriggerType.AfterPrevious);

                ((Sequence)slide.Timeline.MainSequence).AddEffect(chart,
                EffectChartMajorGroupingType.BySeries, 3,
                EffectType.Appear, EffectSubtype.None, EffectTriggerType.AfterPrevious);

                // Write the modified presentation to disk 
                presentation.Save(dataDir + "AnimatingSeries_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:AnimatingSeries
        }
    }
}