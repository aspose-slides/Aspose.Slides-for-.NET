Imports System
Imports Aspose.Slides.Export
Imports Aspose.Slides.Charts
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Text
    Class CustomRotationAngleTextframe
        Public Shared Sub Run()
            ' ExStart:CustomRotationAngleTextframe
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()
            ' Create an instance of Presentation class
            Dim presentation As New Presentation()

            Dim chart As IChart = presentation.Slides(0).Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 500, 300)

            Dim series As IChartSeries = chart.ChartData.Series(0)

            series.Labels.DefaultDataLabelFormat.ShowValue = True
            series.Labels.DefaultDataLabelFormat.TextFormat.TextBlockFormat.RotationAngle = 65

            chart.HasTitle = True
            chart.ChartTitle.AddTextFrameForOverriding("Custom title").TextFrameFormat.RotationAngle = -30

            ' Save Presentation
            presentation.Save(dataDir & Convert.ToString("textframe-rotation_out.pptx"), SaveFormat.Pptx)
            ' ExEnd:CustomRotationAngleTextframe

        End Sub
    End Class
End Namespace