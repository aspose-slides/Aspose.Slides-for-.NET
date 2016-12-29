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
    Public Class DisplayChartLabels
        Public Shared Sub Run()
			'ExStart:DisplayChartLabels
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            Using presentation As New Presentation()
                Dim chart As IChart = presentation.Slides(0).Shapes.AddChart(ChartType.Pie, 50, 50, 500, 400)
                chart.ChartData.Series(0).Labels.DefaultDataLabelFormat.ShowValue = True
                chart.ChartData.Series(0).Labels.DefaultDataLabelFormat.ShowLabelAsDataCallout = True
                chart.ChartData.Series(0).Labels(2).DataLabelFormat.ShowLabelAsDataCallout = False
                presentation.Save(dataDir & Convert.ToString("DisplayChartLabels_out.pptx"), SaveFormat.Pptx)
            End Using
			'ExEnd:DisplayChartLabels

        End Sub
    End Class
End Namespace