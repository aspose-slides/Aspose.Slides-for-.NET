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
    Public Class SecondPlotOptionsforCharts
        Public Shared Sub Run()
			'ExStart:SecondPlotOptionsforCharts
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            ' Create an instance of Presentation class
            Dim presentation As New Presentation()

            ' Add chart on slide
            Dim chart As IChart = presentation.Slides(0).Shapes.AddChart(ChartType.PieOfPie, 50, 50, 500, 400)

            ' Set different properties
            chart.ChartData.Series(0).Labels.DefaultDataLabelFormat.ShowValue = True
            chart.ChartData.Series(0).ParentSeriesGroup.SecondPieSize = 149
            chart.ChartData.Series(0).ParentSeriesGroup.PieSplitBy = Aspose.Slides.Charts.PieSplitType.ByPercentage
            chart.ChartData.Series(0).ParentSeriesGroup.PieSplitPosition = 53

            ' Write presentation to disk
            presentation.Save(dataDir & Convert.ToString("SecondPlotOptionsforCharts_out.pptx"), SaveFormat.Pptx)
			'ExEnd:SecondPlotOptionsforChart
		End Sub
    End Class
End Namespace