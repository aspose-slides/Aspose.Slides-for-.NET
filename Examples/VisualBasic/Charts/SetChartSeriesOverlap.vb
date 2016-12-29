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
    Public Class SetChartSeriesOverlap
        Public Shared Sub Run()
			'ExStart:SetChartSeriesOverlap
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            Using presentation As New Presentation()
                ' Adding chart
                Dim chart As IChart = presentation.Slides(0).Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 600, 400, True)
                Dim series As IChartSeriesCollection = chart.ChartData.Series
                If series(0).Overlap = 0 Then
                    ' Setting series overlap
                    series(0).ParentSeriesGroup.Overlap = -30
                End If

                ' Write the presentation file to disk
                presentation.Save(dataDir & Convert.ToString("SetChartSeriesOverlap_out.pptx"), SaveFormat.Pptx)
            End Using
			'ExEnd:SetChartSeriesOverlap
        End Sub
    End Class
End Namespace