Imports System
Imports Aspose.Slides.Charts
Imports Aspose.Slides.Export
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Charts
    Public Class SetCategoryAxisLabelDistance
        Public Shared Sub Run()
			'ExStart:SetCategoryAxisLabelDistance
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            Dim presentation As New Presentation()

            ' Get reference of the slide
            Dim sld As ISlide = presentation.Slides(0)

            ' Adding a chart on slide
            Dim ch As IChart = sld.Shapes.AddChart(ChartType.ClusteredColumn, 20, 20, 500, 300)

            ' Setting the position of label from axis
            ch.Axes.HorizontalAxis.LabelOffset = 500

            ' Write the presentation file to disk
            presentation.Save(dataDir & Convert.ToString("SetCategoryAxisLabelDistance_out.pptx"), SaveFormat.Pptx)
        End Sub
			'ExEnd:SetCategoryAxisLabelDistance
    End Class
End Namespace