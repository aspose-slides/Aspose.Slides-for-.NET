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
    Public Class SetlegendCustomOptions
        Public Shared Sub Run()
            'ExStart:SetlegendCustomOptions
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            ' Create an instance of Presentation class
            Dim presentation As New Presentation()

            ' Get reference of the slide
            Dim slide As ISlide = presentation.Slides(0)

            ' Add a clustered column chart on the slide
            Dim chart As IChart = slide.Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 500, 500)

            ' Set Legend Properties
            chart.Legend.X = 50 / chart.Width
            chart.Legend.Y = 50 / chart.Height
            chart.Legend.Width = 100 / chart.Width
            chart.Legend.Height = 100 / chart.Height

            ' Write presentation to disk
            presentation.Save(dataDir & Convert.ToString("Legend_out.pptx"), SaveFormat.Pptx)
			'ExEnd:SetlegendCustomOptions
        End Sub
    End Class
End Namespace