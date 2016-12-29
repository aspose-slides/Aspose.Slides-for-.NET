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
    Public Class DisplayPercentageAsLabels
        Public Shared Sub Run()
			'ExStart:DisplayPercentageAsLabels
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            ' Create an instance of Presentation class
            Dim presentation As New Presentation()

            Dim slide As ISlide = presentation.Slides(0)

            Dim chart As IChart = slide.Shapes.AddChart(ChartType.StackedColumn, 20, 20, 400, 400)
            Dim series As IChartSeries = chart.ChartData.Series(0)
            Dim cat As IChartCategory
            Dim total_value As Double = 0.0F

            Dim total_for_Cat As Double() = New Double(chart.ChartData.Categories.Count - 1) {}

            For k As Integer = 0 To chart.ChartData.Categories.Count - 1
                cat = chart.ChartData.Categories(k)

                For i As Integer = 0 To chart.ChartData.Series.Count - 1
                    total_for_Cat(k) = total_for_Cat(k) + Convert.ToDouble(chart.ChartData.Series(i).DataPoints(k).Value.Data)
                Next
            Next

            Dim dataPontPercent As Double = 0.0F

            For x As Integer = 0 To chart.ChartData.Series.Count - 1
                series = chart.ChartData.Series(x)
                series.Labels.DefaultDataLabelFormat.ShowLegendKey = False

                For j As Integer = 0 To series.DataPoints.Count - 1
                    Dim lbl As IDataLabel = series.DataPoints(j).Label
                    dataPontPercent = (Convert.ToDouble(series.DataPoints(j).Value.Data) / total_for_Cat(j)) * 100

                    Dim port As IPortion = New Portion()
                    port.Text = [String].Format("{0:F2} %", dataPontPercent)
                    port.PortionFormat.FontHeight = 8.0F
                    lbl.TextFrameForOverriding.Text = ""
                    Dim para As IParagraph = lbl.TextFrameForOverriding.Paragraphs(0)
                    para.Portions.Add(port)

                    lbl.DataLabelFormat.ShowSeriesName = False
                    lbl.DataLabelFormat.ShowPercentage = False
                    lbl.DataLabelFormat.ShowLegendKey = False
                    lbl.DataLabelFormat.ShowCategoryName = False
                    lbl.DataLabelFormat.ShowBubbleSize = False
                Next
            Next

            ' Save presentation with chart
            presentation.Save(dataDir & Convert.ToString("DisplayPercentageAsLabels_out.pptx"), SaveFormat.Pptx)

			'ExEnd:DisplayPercentageAsLabels
        End Sub
    End Class
End Namespace