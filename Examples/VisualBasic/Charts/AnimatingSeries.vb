Imports System
Imports Aspose.Slides.Charts
Imports Aspose.Slides.Export
Imports Aspose.Slides
Imports Aspose.Slides.Animation

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Charts
    Public Class AnimatingSeries
        Public Shared Sub Run()
			'ExStart:AnimatingSeries
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            ' Instantiate Presentation class that represents a presentation file 
            Using presentation As New Presentation(dataDir & Convert.ToString("ExistingChart.pptx"))
                ' Get reference of the chart object
                Dim slide = TryCast(presentation.Slides(0), Slide)
                Dim shapes = TryCast(slide.Shapes, ShapeCollection)
                Dim chart = TryCast(shapes(0), IChart)

                ' Animate the series
                slide.Timeline.MainSequence.AddEffect(chart, EffectType.Fade, EffectSubtype.None, EffectTriggerType.AfterPrevious)

                DirectCast(slide.Timeline.MainSequence, Sequence).AddEffect(chart, EffectChartMajorGroupingType.BySeries, 0, EffectType.Appear, EffectSubtype.None, EffectTriggerType.AfterPrevious)

                DirectCast(slide.Timeline.MainSequence, Sequence).AddEffect(chart, EffectChartMajorGroupingType.BySeries, 1, EffectType.Appear, EffectSubtype.None, EffectTriggerType.AfterPrevious)

                DirectCast(slide.Timeline.MainSequence, Sequence).AddEffect(chart, EffectChartMajorGroupingType.BySeries, 2, EffectType.Appear, EffectSubtype.None, EffectTriggerType.AfterPrevious)

                DirectCast(slide.Timeline.MainSequence, Sequence).AddEffect(chart, EffectChartMajorGroupingType.BySeries, 3, EffectType.Appear, EffectSubtype.None, EffectTriggerType.AfterPrevious)

                ' Write the modified presentation to disk 
                presentation.Save(dataDir & Convert.ToString("AnimatingSeries_out.pptx"), SaveFormat.Pptx)
            End Using
			'ExEnd:AnimatingSeries
        End Sub
    End Class
End Namespace