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
    Public Class SetAutomaticSeriesFillColor
        Public Shared Sub Run()
			'ExStart:SetAutomaticSeriesFillColor
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            Using presentation As New Presentation()
                ' Creating a clustered column chart
                Dim chart As IChart = presentation.Slides(0).Shapes.AddChart(ChartType.ClusteredColumn, 100, 50, 600, 400)

                ' Setting series fill format to automatic
                For i As Integer = 0 To chart.ChartData.Series.Count - 1
                    chart.ChartData.Series(i).GetAutomaticSeriesColor()
                Next

                ' Write the presentation file to disk
                presentation.Save(dataDir & Convert.ToString("AutoFillSeries_out.pptx"), SaveFormat.Pptx)
            End Using
			'ExEnd:SetAutomaticSeriesFillColor
        End Sub
    End Class
End Namespace