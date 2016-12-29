Imports System
Imports Aspose.Slides.Charts
Imports Aspose.Slides.Export
Imports System.Drawing
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Charts
    Public Class SetMarkerOptions
        Public Shared Sub Run()
			'ExStart:SetMarkerOptions
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            ' Create an instance of Presentation class
            Dim presentation As New Presentation()

            Dim slide As ISlide = presentation.Slides(0)

            ' Creating the default chart
            Dim chart As IChart = slide.Shapes.AddChart(ChartType.LineWithMarkers, 0, 0, 400, 400)

            ' Getting the default chart data worksheet index
            Dim defaultWorksheetIndex As Integer = 0

            ' Getting the chart data worksheet
            Dim fact As IChartDataWorkbook = chart.ChartData.ChartDataWorkbook

            ' Delete demo series
            chart.ChartData.Series.Clear()

            ' Add new series
            chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 1, 1, "Series 1"), chart.Type)

            ' Set the picture
            Dim image1 As System.Drawing.Image = DirectCast(New Bitmap(dataDir & Convert.ToString("aspose-logo.jpg")), System.Drawing.Image)
            Dim imgx1 As IPPImage = presentation.Images.AddImage(image1)

            ' Set the picture
            Dim image2 As System.Drawing.Image = DirectCast(New Bitmap(dataDir & Convert.ToString("Tulips.jpg")), System.Drawing.Image)
            Dim imgx2 As IPPImage = presentation.Images.AddImage(image2)

            ' Take first chart series
            Dim series As IChartSeries = chart.ChartData.Series(0)

            ' Add new point (1:3) there.
            Dim point As IChartDataPoint = series.DataPoints.AddDataPointForLineSeries(fact.GetCell(defaultWorksheetIndex, 1, 1, CDbl(4.5)))
            point.Marker.Format.Fill.FillType = FillType.Picture
            point.Marker.Format.Fill.PictureFillFormat.Picture.Image = imgx1

            point = series.DataPoints.AddDataPointForLineSeries(fact.GetCell(defaultWorksheetIndex, 2, 1, CDbl(2.5)))
            point.Marker.Format.Fill.FillType = FillType.Picture
            point.Marker.Format.Fill.PictureFillFormat.Picture.Image = imgx2

            point = series.DataPoints.AddDataPointForLineSeries(fact.GetCell(defaultWorksheetIndex, 3, 1, CDbl(3.5)))
            point.Marker.Format.Fill.FillType = FillType.Picture
            point.Marker.Format.Fill.PictureFillFormat.Picture.Image = imgx1

            point = series.DataPoints.AddDataPointForLineSeries(fact.GetCell(defaultWorksheetIndex, 4, 1, CDbl(4.5)))
            point.Marker.Format.Fill.FillType = FillType.Picture
            point.Marker.Format.Fill.PictureFillFormat.Picture.Image = imgx2

            ' Changing the chart series marker
            series.Marker.Size = 15

            ' Write presentation to disk
            presentation.Save(dataDir & Convert.ToString("MarkOptions_out.pptx"), SaveFormat.Pptx)
			'ExEnd:SetMarkerOptions
		End Sub
    End Class
End Namespace